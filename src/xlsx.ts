import { FormatError, capped, withLimits, type Limits } from "./limits.js";
import { parseRef, refOf, sortedRefs, type Cell, type CellValue, type ReadResult, type Sheet, type Workbook, type WriteResult } from "./model.js";
import { REL_NS, XML_HEADER, contentTypesXml, readOoxmlPackage, relsXml } from "./ooxml.js";
import { MAX_FORMULA, checkWorkbook, hasContent, isBooleanLiteral, isoToSerial } from "./spreadsheet.js";
import { childOf, elementsOf, textOf, xmlEscape as esc, type XmlElement } from "./xml.js";
import { writeZip } from "./zip.js";

const MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const SML = "application/vnd.openxmlformats-officedocument.spreadsheetml";
const BUILTIN_FORMATS: Record<string, number> = { "0": 1, "0.00": 2, "#,##0": 3, "#,##0.00": 4, "0%": 9, "0.00%": 10 };

// ---------- write ----------

function cellXml(ref: string, cell: Cell, xf: (format: string) => number): string {
  const style = cell.format ? xf(cell.format) : 0;
  const attrs = `r="${ref}"${style ? ` s="${style}"` : ""}`;
  const f = cell.formula !== undefined ? `<f>${esc(cell.formula.replace(/^=/, ""))}</f>` : "";
  const { value } = cell;
  if (typeof value === "string") {
    return f ? `<c ${attrs} t="str">${f}<v>${esc(value)}</v></c>` : `<c ${attrs} t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
  }
  if (typeof value === "boolean") return `<c ${attrs} t="b">${f}<v>${value ? 1 : 0}</v></c>`;
  if (typeof value === "number") return `<c ${attrs}>${f}<v>${value}</v></c>`;
  return f ? `<c ${attrs}>${f}</c>` : `<c ${attrs}/>`;
}

/** Write a workbook as an Excel file (.xlsx). Strings are inline; cached formula results are kept. */
export function writeXlsx(book: Workbook, limits?: Partial<Limits>): WriteResult {
  const caps = withLimits(limits);
  checkWorkbook(book, caps);
  const formats: string[] = [];
  const formatIds = new Map<string, number>();
  const xf = (format: string): number => {
    let id = formatIds.get(format);
    if (id === undefined) {
      id = formats.push(format);
      formatIds.set(format, id);
    }
    return id;
  };
  const chargeTotal = capped(caps.maxTotalBytes, `the export passes ${caps.maxTotalBytes} bytes while it is written`);
  const sheets = book.sheets.map((sheet) => {
    // A running cap, so a model with huge cells stops before the whole sheet exists as one string.
    const chargeSheet = capped(caps.maxEntryBytes, `the export passes ${caps.maxEntryBytes} bytes in one sheet while it is written`);
    const rows = new Map<number, string[]>();
    for (const ref of sortedRefs(sheet)) {
      if (!hasContent(sheet.cells[ref])) continue;
      const row = parseRef(ref)!.row;
      const xml = cellXml(ref, sheet.cells[ref], xf);
      chargeSheet(xml.length);
      chargeTotal(xml.length);
      const inRow = rows.get(row);
      if (inRow) inRow.push(xml);
      else rows.set(row, [xml]);
    }
    const data = [...rows]
      .map(([row, cells]) => {
        const wrapper = `<row r="${row}"></row>`.length;
        chargeSheet(wrapper);
        chargeTotal(wrapper);
        return `<row r="${row}">${cells.join("")}</row>`;
      })
      .join("");
    return `${XML_HEADER}<worksheet xmlns="${MAIN}"><sheetData>${data}</sheetData></worksheet>`;
  });
  let custom = 164;
  const xfs = formats.map((format) => {
    const id = BUILTIN_FORMATS[format] ?? custom++;
    return { id, format, builtin: format in BUILTIN_FORMATS };
  });
  const numFmts = xfs.filter((x) => !x.builtin);
  const styles =
    `${XML_HEADER}<styleSheet xmlns="${MAIN}">` +
    (numFmts.length ? `<numFmts count="${numFmts.length}">${numFmts.map((x) => `<numFmt numFmtId="${x.id}" formatCode="${esc(x.format)}"/>`).join("")}</numFmts>` : "") +
    `<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>` +
    `<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="${xfs.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>` +
    xfs.map((x) => `<xf numFmtId="${x.id}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join("") +
    `</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
  const workbook =
    `${XML_HEADER}<workbook xmlns="${MAIN}" xmlns:r="${REL_NS}"><sheets>` +
    book.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") +
    `</sheets></workbook>`;
  const sheetPaths = book.sheets.map((_, i) => `xl/worksheets/sheet${i + 1}.xml`);
  const bytes = writeZip(
    [
      {
        name: "[Content_Types].xml",
        data: contentTypesXml(
          { rels: "application/vnd.openxmlformats-package.relationships+xml", xml: "application/xml" },
          { "xl/workbook.xml": `${SML}.sheet.main+xml`, "xl/styles.xml": `${SML}.styles+xml`, ...Object.fromEntries(sheetPaths.map((p) => [p, `${SML}.worksheet+xml`])) },
        ),
      },
      { name: "_rels/.rels", data: relsXml([{ id: "rId1", type: "officeDocument", target: "xl/workbook.xml" }]) },
      { name: "xl/workbook.xml", data: workbook },
      {
        name: "xl/_rels/workbook.xml.rels",
        data: relsXml([...sheetPaths.map((p, i) => ({ id: `rId${i + 1}`, type: "worksheet", target: p.slice(3) })), { id: `rId${sheetPaths.length + 1}`, type: "styles", target: "styles.xml" }]),
      },
      { name: "xl/styles.xml", data: styles },
      ...sheetPaths.map((name, i) => ({ name, data: sheets[i] })),
    ],
    caps,
  );
  return { bytes, warnings: [] };
}

// ---------- read ----------

const BUILTIN_BY_ID: Record<number, string> = { 1: "0", 2: "0.00", 3: "#,##0", 4: "#,##0.00", 9: "0%", 10: "0.00%", 14: "yyyy-mm-dd" };

/** Known spellings of one format become the model's code. Any other code stays as written. */
function canonicalFormat(code: string): string | undefined {
  const plain = code.replace(/\\(.)/g, "$1").replace(/"/g, "");
  // LibreOffice gives a boolean cell the display format "TRUE";"TRUE";"FALSE".
  if (/^general$/i.test(plain) || /^true;true;false$/i.test(plain)) return undefined;
  return /^yyyy-mm-dd$/i.test(plain) ? "yyyy-mm-dd" : code;
}

/** The text of a shared or inline string: every `t`, and none of the phonetic hints. */
const stringOf = (si: XmlElement): string =>
  elementsOf(si)
    .filter((k) => k.name !== "rPh" && k.name !== "phoneticPr")
    .map((k) => (k.name === "t" ? textOf(k) : k.name === "r" ? textOf(childOf(k, "t")) : ""))
    .join("");

/** Read an Excel file (.xlsx). Content with no place in the model is listed in `warnings`. */
export function readXlsx(bytes: Uint8Array, limits?: Partial<Limits>): ReadResult<Workbook> {
  const caps = withLimits(limits);
  const pkg = readOoxmlPackage(bytes, caps, "XLSX", SML);
  const warnings = new Set<string>();
  const warn = (message: string) => warnings.add(message);
  const workbook = pkg.parse(pkg.mainPart);
  if (childOf(workbook, "workbookPr")?.attrs.date1904 === "1") warn("the 1904 date system: dates read 1462 days early");
  const workbookRels = pkg.rels(pkg.mainPart);
  const partOf = (kind: string): string | undefined => [...workbookRels.values()].find((r) => r.kind === kind && !r.external)?.target;

  const shared: string[] = [];
  const sharedPart = partOf("sharedStrings");
  if (sharedPart && pkg.files[sharedPart]) for (const si of elementsOf(pkg.parse(sharedPart))) shared.push(stringOf(si));

  const formatOfStyle: (string | undefined)[] = [];
  const stylesPart = partOf("styles");
  if (stylesPart && pkg.files[stylesPart]) {
    const styles = pkg.parse(stylesPart);
    const codes = new Map(elementsOf(childOf(styles, "numFmts")).map((n) => [n.attrs.numFmtId, n.attrs.formatCode ?? ""]));
    for (const xf of elementsOf(childOf(styles, "cellXfs"))) {
      const id = Number(xf.attrs.numFmtId ?? 0);
      if (codes.has(String(id))) formatOfStyle.push(canonicalFormat(codes.get(String(id))!));
      else {
        if (id !== 0 && !BUILTIN_BY_ID[id]) warn(`the built-in number format ${id} was not read`);
        formatOfStyle.push(BUILTIN_BY_ID[id]);
      }
    }
  }

  let cellCount = 0;
  // A shared string is one string in memory but fills every cell that points to it, so each use counts.
  const chargeText = capped(caps.maxTextChars, `the workbook has more than ${caps.maxTextChars} characters of cell text`);
  const readParts = new Set<string>();
  const sheets: Sheet[] = [];
  for (const entry of elementsOf(childOf(workbook, "sheets"))) {
    const rel = workbookRels.get(entry.attrs["r:id"] ?? "");
    if (!rel || rel.external || rel.kind !== "worksheet") {
      warn(`the sheet "${entry.attrs.name}" is not a worksheet: not imported`);
      continue;
    }
    // A part that two sheet entries name would be parsed and copied twice.
    if (readParts.has(rel.target)) {
      warn(`the sheet "${entry.attrs.name}" shares its worksheet part with an earlier sheet: not imported`);
      continue;
    }
    readParts.add(rel.target);
    const sheet = pkg.parse(rel.target);
    const cells: Record<string, Cell> = {};
    if (childOf(sheet, "mergeCells")) warn("merged cells were read as single cells");
    let nextRow = 0;
    for (const row of elementsOf(childOf(sheet, "sheetData"))) {
      nextRow = Number(row.attrs.r) || nextRow + 1;
      let nextColumn = 0;
      for (const c of elementsOf(row)) {
        const at = c.attrs.r ? parseRef(c.attrs.r) : parseRef(refOf(nextColumn + 1, nextRow));
        if (!at) throw new FormatError("INVALID_FILE", "a cell reference is not on the sheet grid");
        nextColumn = at.column;
        const type = c.attrs.t;
        const raw = textOf(childOf(c, "v"));
        let value: CellValue | undefined;
        if (type === "s") {
          const text = shared[Number(raw)];
          if (text === undefined) warn("a cell points to a missing shared string");
          else value = text;
        } else if (type === "inlineStr") value = stringOf(childOf(c, "is") ?? c);
        else if (type === "str") value = raw;
        else if (type === "b") value = raw === "1";
        else if (type === "e") {
          value = raw;
          warn("an error value was read as text");
        } else if (type === "d") value = isoToSerial(raw);
        else if (childOf(c, "v") && raw.trim() !== "" && Number.isFinite(Number(raw))) value = Number(raw);
        const f = childOf(c, "f");
        let formula: string | undefined;
        if (f) {
          if (f.attrs.t === "shared" && !textOf(f)) warn("a shared formula was read as its value");
          else formula = textOf(f).replace(/^=/, "");
          if (formula !== undefined && formula.length > MAX_FORMULA) throw new FormatError("LIMIT_EXCEEDED", `a formula is longer than ${MAX_FORMULA} characters`);
          if (formula !== undefined && isBooleanLiteral(formula, value)) formula = undefined;
        }
        const format = formatOfStyle[Number(c.attrs.s ?? 0)];
        if (value === undefined && formula === undefined && !(format && c.attrs.s)) continue;
        chargeText((typeof value === "string" ? value.length : 0) + (formula?.length ?? 0));
        if (++cellCount > caps.maxCells) throw new FormatError("LIMIT_EXCEEDED", `the workbook has more than ${caps.maxCells} cells`);
        cells[refOf(at.column, at.row)] = { ...(value !== undefined ? { value } : {}), ...(formula !== undefined ? { formula } : {}), ...(format ? { format } : {}) };
      }
    }
    sheets.push({ name: entry.attrs.name ?? `Sheet${sheets.length + 1}`, cells });
  }
  if (!sheets.length) throw new FormatError("INVALID_FILE", "the XLSX file has no sheet");
  return { model: { sheets }, warnings: [...warnings] };
}
