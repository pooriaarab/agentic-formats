import { toOpenFormula, fromOpenFormula } from "./formula.js";
import { FormatError, capped, withLimits, type Limits } from "./limits.js";
import { parseRef, refOf, sortedRefs, type Cell, type CellValue, type ReadResult, type Sheet, type Workbook, type WriteResult } from "./model.js";
import { ODF_NS_ATTRS, XML_HEADER, readOdfPackage, spaced, writeOdfPackage } from "./odf.js";
import { MAX_FORMULA, ODS_FORMAT, checkWorkbook, hasContent, isBooleanLiteral, isDateFormat, isoToSerial, serialToIso } from "./spreadsheet.js";
import { childOf, elementsOf, xmlEscape as esc, type XmlElement } from "./xml.js";

export const ODS_MIME = "application/vnd.oasis.opendocument.spreadsheet";
const READABLE_MIMES = [ODS_MIME, `${ODS_MIME}-template`];
const MAX_COLUMN = 16384;
const MAX_FORMAT_WARNINGS = 100;
/** LibreOffice reads the `of:` in a formula as an XML namespace prefix, so the root must declare it. */
const FORMULA_NS = 'xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"';
const STYLES = `${XML_HEADER}<office:document-styles ${ODF_NS_ATTRS} office:version="1.3"><office:styles><style:style style:name="Default" style:family="table-cell"/></office:styles></office:document-styles>`;

// ---------- write ----------

const decimalsOf = (format: string): number => (format.split(".")[1] ?? "").replace("%", "").length;

function dataStyleXml(format: string, name: string): string {
  if (isDateFormat(format)) {
    return `<number:date-style style:name="${name}"><number:year number:style="long"/><number:text>-</number:text><number:month number:style="long"/><number:text>-</number:text><number:day number:style="long"/></number:date-style>`;
  }
  const number = `<number:number number:decimal-places="${decimalsOf(format)}" number:min-integer-digits="1"${format.startsWith("#,##0") ? ' number:grouping="true"' : ""}/>`;
  return format.endsWith("%")
    ? `<number:percentage-style style:name="${name}">${number}<number:text>%</number:text></number:percentage-style>`
    : `<number:number-style style:name="${name}">${number}</number:number-style>`;
}

function cellXml(cell: Cell, styleOf: (format: string) => string | undefined): string {
  const attrs: string[] = [];
  const style = cell.format ? styleOf(cell.format) : undefined;
  if (style) attrs.push(`table:style-name="${style}"`);
  if (cell.formula !== undefined) attrs.push(`table:formula="of:=${esc(toOpenFormula(cell.formula.replace(/^=/, "")))}"`);
  const { value } = cell;
  let text = "";
  if (typeof value === "number") {
    const iso = style && isDateFormat(cell.format) ? serialToIso(value) : undefined;
    if (iso) attrs.push(`office:value-type="date" office:date-value="${iso}"`);
    else attrs.push(`office:value-type="${cell.format?.endsWith("%") && style ? "percentage" : "float"}" office:value="${value}"`);
    text = iso ?? String(value);
  } else if (typeof value === "boolean") {
    attrs.push(`office:value-type="boolean" office:boolean-value="${value}"`);
    text = value ? "TRUE" : "FALSE";
  } else if (typeof value === "string") {
    attrs.push(`office:value-type="string"${cell.formula !== undefined ? ` office:string-value="${esc(value)}"` : ""}`);
    text = value;
  }
  return `<table:table-cell ${attrs.join(" ")}>${text !== "" || value !== undefined ? `<text:p>${spaced(text)}</text:p>` : ""}</table:table-cell>`;
}

function sheetXml(sheet: Sheet, styleOf: (format: string) => string | undefined, charge: (bytes: number) => void): string {
  const byRow = new Map<number, [number, Cell][]>();
  let columns = 1;
  for (const ref of sortedRefs(sheet)) {
    const cell = sheet.cells[ref];
    if (!hasContent(cell)) continue;
    const at = parseRef(ref)!;
    columns = Math.max(columns, at.column);
    const inRow = byRow.get(at.row);
    if (inRow) inRow.push([at.column, cell]);
    else byRow.set(at.row, [[at.column, cell]]);
  }
  let rows = "";
  let row = 0;
  for (const [number, cells] of byRow) {
    if (number > row + 1) {
      const gap = `<table:table-row table:number-rows-repeated="${number - row - 1}"><table:table-cell/></table:table-row>`;
      charge(gap.length);
      rows += gap;
    }
    let column = 0;
    const inRow = cells
      .map(([at, cell]) => {
        const gap = at > column + 1 ? `<table:table-cell table:number-columns-repeated="${at - column - 1}"/>` : "";
        column = at;
        const xml = gap + cellXml(cell, styleOf);
        charge(xml.length);
        return xml;
      })
      .join("");
    charge("<table:table-row></table:table-row>".length);
    rows += `<table:table-row>${inRow}</table:table-row>`;
    row = number;
  }
  return `<table:table table:name="${esc(sheet.name)}"><table:table-column table:number-columns-repeated="${columns}"/>${rows || "<table:table-row><table:table-cell/></table:table-row>"}</table:table>`;
}

/** Write a workbook as an OpenDocument spreadsheet (.ods). */
export function writeOds(book: Workbook, limits?: Partial<Limits>): WriteResult {
  const caps = withLimits(limits);
  checkWorkbook(book, caps);
  const warnings = new Set<string>();
  const formats = new Map<string, string>();
  const styleOf = (format: string): string | undefined => {
    if (!ODS_FORMAT.test(format)) {
      // A hostile workbook can carry a different format code in every cell; keep the list short.
      if (warnings.size < MAX_FORMAT_WARNINGS) warnings.add(`the number format "${format}" is not written in ODS`);
      else warnings.add("more number formats are not written in ODS");
      return undefined;
    }
    if (!formats.has(format)) formats.set(format, `${formats.size + 1}`);
    return `ce${formats.get(format)}`;
  };
  // A running cap, so a model with huge cells stops before the content exists as one string.
  const charge = capped(caps.maxEntryBytes, `the export passes ${caps.maxEntryBytes} bytes while it is written`);
  const tables = book.sheets.map((sheet) => sheetXml(sheet, styleOf, charge)).join("");
  const styles = [...formats].map(([format, n]) => `${dataStyleXml(format, `N${n}`)}<style:style style:name="ce${n}" style:family="table-cell" style:parent-style-name="Default" style:data-style-name="N${n}"/>`).join("");
  const content =
    `${XML_HEADER}<office:document-content ${ODF_NS_ATTRS} ${FORMULA_NS} office:version="1.3"><office:automatic-styles>${styles}</office:automatic-styles>` +
    `<office:body><office:spreadsheet>${tables}</office:spreadsheet></office:body></office:document-content>`;
  return { bytes: writeOdfPackage({ mime: ODS_MIME, content, styles: STYLES, images: [] }, caps), warnings: [...warnings] };
}

// ---------- read ----------

/** The number format of an ODF data style, or undefined when the style is outside the formats we keep. */
function formatOfDataStyle(style: XmlElement, warn: (message: string) => void): string | undefined {
  const kids = elementsOf(style).filter((k) => !(k.name === "number:text" && !textOfKid(k).trim()) && k.name !== "style:text-properties" && k.name !== "style:map");
  const number = kids.find((k) => k.name === "number:number");
  const decimals = Math.max(0, Math.min(30, Number.parseInt(number?.attrs["number:decimal-places"] ?? "0", 10) || 0));
  const zeros = decimals ? `.${"0".repeat(decimals)}` : "";
  switch (style.name) {
    case "number:number-style":
      return kids.length === 1 && number ? `${number.attrs["number:grouping"] === "true" ? "#,##0" : "0"}${zeros}` : undefined;
    case "number:percentage-style":
      return kids.length === 2 && number && textOfKid(kids[1]) === "%" ? `0${zeros}%` : undefined;
    case "number:date-style":
      if (kids.map((k) => k.name).filter((n) => n !== "number:text").join() !== "number:year,number:month,number:day") warn("a date format was read as yyyy-mm-dd");
      return "yyyy-mm-dd";
    default:
      return undefined;
  }
}

const textOfKid = (el: XmlElement): string => el.children.map((c) => (typeof c === "string" ? c : textOfKid(c))).join("");

/** A `<text:s>` element stands for this many spaces at most. */
const MAX_SPACES = 1000;

/** The text of a cell. `spaces` is told how many spaces each `<text:s>` stands for, before the string that holds them is made. */
function cellText(cell: XmlElement, spaces: (count: number, asked: number) => void): string {
  const walk = (el: XmlElement): string =>
    el.children
      .map((c) => {
        if (typeof c === "string") return c.replace(/[ \t\r\n]+/g, " ");
        if (c.name === "text:s") {
          const asked = Number(c.attrs["text:c"] ?? 1) || 1;
          const count = Math.max(1, Math.min(MAX_SPACES, asked));
          spaces(count, asked);
          return "\u0001".repeat(count);
        }
        if (c.name === "text:tab") return "\u0002";
        if (c.name === "text:line-break") return "\n";
        return c.name === "text:p" || c.name === "text:h" ? "" : walk(c);
      })
      .join("");
  return elementsOf(cell)
    .filter((p) => p.name === "text:p" || p.name === "text:h")
    .map((p) => walk(p).trim().replaceAll("\u0001", " ").replaceAll("\u0002", "\t"))
    .join("\n");
}

/** Read an OpenDocument spreadsheet (.ods). Content with no place in the model is listed in `warnings`. */
export function readOds(bytes: Uint8Array, limits?: Partial<Limits>): ReadResult<Workbook> {
  const caps = withLimits(limits);
  const { content, styles: stylesPart } = readOdfPackage(bytes, caps, "ODS", READABLE_MIMES);
  const body = childOf(childOf(content, "office:body"), "office:spreadsheet");
  if (!body) throw new FormatError("INVALID_FILE", "the ODS file has no spreadsheet body");
  const warnings = new Set<string>();
  const warn = (message: string) => warnings.add(message);

  const dataFormats = new Map<string, string | undefined>();
  const cellStyles = new Map<string, { parent?: string; data?: string }>();
  const collect = (el: XmlElement): void => {
    const name = el.attrs["style:name"];
    if (name && /^number:(number|percentage|date)-style$/.test(el.name)) dataFormats.set(name, formatOfDataStyle(el, warn));
    else if (name && el.name === "style:style" && el.attrs["style:family"] === "table-cell") cellStyles.set(name, { parent: el.attrs["style:parent-style-name"], data: el.attrs["style:data-style-name"] });
    elementsOf(el).forEach(collect);
  };
  collect(content);
  if (stylesPart) collect(stylesPart);
  const formatOfStyle = (name: string | undefined): string | undefined => {
    const seen = new Set<string>();
    for (let n = name, hops = 0; n && !seen.has(n); n = cellStyles.get(n)?.parent, hops++) {
      seen.add(n);
      if (hops >= caps.maxStyleDepth) throw new FormatError("LIMIT_EXCEEDED", `a style has more than ${caps.maxStyleDepth} parents`);
      const data = cellStyles.get(n)?.data;
      if (data !== undefined) return dataFormats.get(data);
    }
    return undefined;
  };

  let cellCount = 0;
  // A repeated cell holds the same text in every copy, so each copy counts.
  const chargeText = capped(caps.maxTextChars, `the workbook has more than ${caps.maxTextChars} characters of cell text`);
  const chargeSpaces = capped(caps.maxTextChars, `the workbook has more than ${caps.maxTextChars} spaces in <text:s> elements`);
  const spaces = (count: number, asked: number) => {
    if (asked > MAX_SPACES) warn(`a run of more than ${MAX_SPACES} spaces was shortened`);
    chargeSpaces(count);
  };
  const count = (n: number) => {
    cellCount += n;
    if (cellCount > caps.maxCells) throw new FormatError("LIMIT_EXCEEDED", `the workbook has more than ${caps.maxCells} cells`);
  };
  const repeatOf = (el: XmlElement, attr: string): number => {
    const n = Number.parseInt(el.attrs[attr] ?? "1", 10);
    return Number.isFinite(n) && n > 1 ? n : 1;
  };

  function cellOf(c: XmlElement, columnStyle: string | undefined): Cell | undefined {
    const type = c.attrs["office:value-type"];
    let value: CellValue | undefined;
    let format = formatOfStyle(c.attrs["table:style-name"] ?? columnStyle);
    if (type === "float" || type === "currency" || type === "percentage") {
      const n = Number(c.attrs["office:value"]);
      if (Number.isFinite(n)) value = n;
      if (type === "currency") warn("a currency cell was read as a number");
    } else if (type === "string") value = c.attrs["office:string-value"] ?? cellText(c, spaces);
    else if (type === "boolean") value = c.attrs["office:boolean-value"] === "true";
    else if (type === "date") {
      value = isoToSerial(c.attrs["office:date-value"] ?? "");
      format ??= "yyyy-mm-dd";
    } else if (type === "time") {
      const t = /^PT(\d+)H(\d+)M(\d+(?:\.\d+)?)S$/.exec(c.attrs["office:time-value"] ?? "");
      if (t) value = (Number(t[1]) * 3600 + Number(t[2]) * 60 + Number(t[3])) / 86400;
      warn("a time cell was read as a fraction of a day");
    }
    let formula: string | undefined;
    const raw = c.attrs["table:formula"];
    if (raw !== undefined) {
      if (raw.length > MAX_FORMULA) throw new FormatError("LIMIT_EXCEEDED", `a formula is longer than ${MAX_FORMULA} characters`);
      chargeText(raw.length);
      formula = raw.startsWith("of:=") ? fromOpenFormula(raw.slice(4)) : undefined;
      if (formula === undefined) warn("a formula that is not in OpenFormula A1 form was read as its value");
      else if (isBooleanLiteral(formula, value)) formula = undefined;
    }
    if (value === undefined && formula === undefined && !(format && c.attrs["table:style-name"])) return undefined;
    return { ...(value !== undefined ? { value } : {}), ...(formula !== undefined ? { formula } : {}), ...(format ? { format } : {}) };
  }

  const sheets: Sheet[] = [];
  for (const table of elementsOf(body).filter((t) => t.name === "table:table")) {
    const cells: Record<string, Cell> = {};
    const columnStyles: (string | undefined)[] = [];
    let row = 0;
    const walk = (parent: XmlElement): void => {
      for (const k of elementsOf(parent)) {
        if (k.name === "table:table-column") {
          const style = k.attrs["table:default-cell-style-name"];
          for (let i = repeatOf(k, "table:number-columns-repeated"); i > 0 && columnStyles.length < MAX_COLUMN; i--) columnStyles.push(style);
        } else if (k.name === "table:table-row") {
          const rowRepeat = repeatOf(k, "table:number-rows-repeated");
          const found: [number, Cell][] = [];
          let column = 0;
          for (const c of elementsOf(k)) {
            if (c.name !== "table:table-cell" && c.name !== "table:covered-table-cell") continue;
            const repeat = repeatOf(c, "table:number-columns-repeated");
            const cell = cellOf(c, columnStyles[column]);
            if (c.name === "table:covered-table-cell" || c.attrs["table:number-columns-spanned"] || c.attrs["table:number-rows-spanned"]) warn("merged cells were read as single cells");
            if (cell) {
              count(repeat * rowRepeat);
              chargeText(((typeof cell.value === "string" ? cell.value.length : 0) + (cell.formula?.length ?? 0)) * repeat * rowRepeat);
              if (column + repeat > MAX_COLUMN) throw new FormatError("LIMIT_EXCEEDED", "a cell lies outside the sheet grid");
              for (let i = 0; i < repeat; i++) found.push([column + i + 1, cell]);
            }
            column += repeat;
          }
          if (found.length && row + rowRepeat > 1048576) throw new FormatError("LIMIT_EXCEEDED", "a cell lies outside the sheet grid");
          for (let r = 1; r <= rowRepeat && found.length; r++) for (const [at, cell] of found) cells[refOf(at, row + r)] = { ...cell };
          row += rowRepeat;
        } else if (/^table:(table-header-rows|table-rows|table-row-group)$/.test(k.name)) walk(k);
      }
    };
    walk(table);
    sheets.push({ name: table.attrs["table:name"] ?? `Sheet${sheets.length + 1}`, cells });
  }
  if (!sheets.length) throw new FormatError("INVALID_FILE", "the ODS file has no sheet");
  return { model: { sheets }, warnings: [...warnings] };
}
