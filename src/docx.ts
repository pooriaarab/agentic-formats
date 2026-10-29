import { frameSize, sniffImage } from "./image.js";
import { FormatError, withLimits, type Limits } from "./limits.js";
import {
  MARKS,
  normalizeInline,
  normalizeTextDoc,
  type Block,
  type Footnote,
  type Heading,
  type Inline,
  type InlineImage,
  type List,
  type Mark,
  type ReadResult,
  type Table,
  type TextDoc,
  type TextRun,
  type WriteResult,
} from "./model.js";
import { REL_NS, XML_HEADER, contentTypesXml, readOoxmlPackage, relsXml } from "./ooxml.js";
import { childOf, elementsOf, safeHref, textOf, xmlEscape as esc, type XmlElement } from "./xml.js";
import { writeZip, type ZipEntry } from "./zip.js";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WML = "application/vnd.openxmlformats-officedocument.wordprocessingml";
const EMU = 12700;
const NS = `xmlns:w="${W}" xmlns:r="${REL_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"`;
const CODE_FONT = '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>';
const MARK_XML: Record<Mark, string> = {
  code: CODE_FONT,
  bold: "<w:b/>",
  italic: "<w:i/>",
  strike: "<w:strike/>",
  underline: '<w:u w:val="single"/>',
  superscript: '<w:vertAlign w:val="superscript"/>',
  subscript: '<w:vertAlign w:val="subscript"/>',
};
// The order the schema asks for inside a run's properties.
const RPR_ORDER: Mark[] = ["code", "bold", "italic", "strike", "underline", "superscript", "subscript"];

// ---------- write ----------

const style = (id: string, name: string, extra: string, type = "paragraph") =>
  `<w:style w:type="${type}" w:styleId="${id}"><w:name w:val="${name}"/>${extra}</w:style>`;
const HEADING_SIZES = [48, 36, 28, 24, 22, 20];
const STYLES =
  `${XML_HEADER}<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>` +
  `<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="259" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  `<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>` +
  `<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
  HEADING_SIZES.map((size, i) =>
    style(`Heading${i + 1}`, `heading ${i + 1}`, `<w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="60"/><w:outlineLvl w:val="${i}"/></w:pPr><w:rPr><w:b/><w:sz w:val="${size}"/></w:rPr>`),
  ).join("") +
  style("ListParagraph", "List Paragraph", '<w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr>') +
  style("FootnoteText", "footnote text", '<w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:sz w:val="20"/></w:rPr>') +
  style("FootnoteReference", "footnote reference", '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>', "character") +
  style("Hyperlink", "Hyperlink", '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>', "character") +
  style("TableGrid", "Table Grid", `<w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join("")}</w:tblBorders></w:tblPr>`, "table") +
  `</w:styles>`;

const level = (i: number, ordered: boolean) =>
  `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${ordered ? "decimal" : "bullet"}"/><w:lvlText w:val="${ordered ? `%${i + 1}.` : "•"}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
const abstractNum = (id: number, ordered: boolean) =>
  `<w:abstractNum w:abstractNumId="${id}"><w:multiLevelType w:val="hybridMultilevel"/>${Array.from({ length: 9 }, (_, i) => level(i, ordered)).join("")}</w:abstractNum>`;

interface Out {
  /** Hyperlinks and images, as relationships of the document part. */
  rels: { id: string; type: string; target: string; external?: boolean }[];
  media: ZipEntry[];
  nums: boolean[]; // one entry per list; true when ordered
  notes: string[];
  warnings: Set<string>;
}

const nextId = (out: Out) => `rId${out.rels.length + 4}`; // rId1 to rId3 belong to styles, numbering, and footnotes

function textPieces(text: string): string {
  return text
    .split(/(\t|\n)/)
    .map((piece) => (piece === "\t" ? "<w:tab/>" : piece === "\n" ? "<w:br/>" : piece ? `<w:t xml:space="preserve">${esc(piece)}</w:t>` : ""))
    .join("");
}

function imageXml(node: InlineImage, out: Out): string {
  const sniffed = sniffImage(node.data);
  if (!sniffed) {
    out.warnings.add("an image is not a PNG, JPEG, or GIF: not written");
    return "";
  }
  const n = out.media.length + 1;
  const id = nextId(out);
  out.media.push({ name: `word/media/image${n}.${sniffed.ext}`, data: node.data });
  out.rels.push({ id, type: "image", target: `media/image${n}.${sniffed.ext}` });
  const [w, h] = frameSize(node.width, node.height, sniffed);
  const [cx, cy] = [Math.round(w * EMU), Math.round(h * EMU)];
  return (
    `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${n}" name="Picture ${n}" descr="${esc(node.alt ?? "")}"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${n}" name="image${n}.${sniffed.ext}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
  );
}

function runXml(run: TextRun, out: Out): string {
  const href = run.href === undefined ? undefined : safeHref(run.href);
  if (run.href !== undefined && !href) out.warnings.add("a link with an unsafe target was written as plain text");
  const props = (href ? '<w:rStyle w:val="Hyperlink"/>' : "") + RPR_ORDER.filter((m) => run.marks?.includes(m)).map((m) => MARK_XML[m]).join("");
  const xml = `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}${textPieces(run.text)}</w:r>`;
  if (!href) return xml;
  if (href.startsWith("#")) return `<w:hyperlink w:anchor="${esc(href.slice(1))}">${xml}</w:hyperlink>`;
  const id = nextId(out);
  out.rels.push({ id, type: "hyperlink", target: href, external: true });
  return `<w:hyperlink r:id="${id}">${xml}</w:hyperlink>`;
}

function inlineXml(nodes: Inline[], out: Out): string {
  return nodes
    .map((node) => {
      switch (node.type) {
        case "text":
          return runXml(node, out);
        case "break":
          return "<w:r><w:br/></w:r>";
        case "image":
          return imageXml(node, out);
        case "footnote": {
          const id = out.notes.length + 1;
          out.notes.push(
            `<w:footnote w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>${inlineXml(node.content, out)}</w:p></w:footnote>`,
          );
          return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
        }
      }
    })
    .join("");
}

function listXml(list: List, depth: number, out: Out): string {
  if (depth > 8) out.warnings.add("a list nests deeper than nine levels: the deepest levels were flattened");
  const level = Math.min(depth, 8);
  const numId = out.nums.push(list.ordered);
  return list.items
    .map(
      (item) =>
        `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${inlineXml(item.content, out)}</w:p>` +
        (item.sublist ? listXml(item.sublist, depth + 1, out) : ""),
    )
    .join("");
}

function tableXml(table: Table, out: Out): string {
  const columns = Math.max(1, ...table.rows.map((r) => r.length));
  const width = Math.floor(9000 / columns);
  const row = (cells: Block[][], header: boolean) =>
    `<w:tr>${header ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${Array.from({ length: columns }, (_, i) => {
      const blocks = (cells[i] ?? []).map((b) => blockXml(b, out)).join("");
      const last = cells[i]?.[cells[i].length - 1];
      return `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${blocks}${!last || last.type === "table" ? "<w:p/>" : ""}</w:tc>`;
    }).join("")}</w:tr>`;
  return (
    `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${width}"/>`.repeat(columns)}</w:tblGrid>` +
    `${table.rows.map((cells, i) => row(cells, !!table.header && i === 0)).join("")}</w:tbl>`
  );
}

function blockXml(block: Block, out: Out, depth = 0): string {
  switch (block.type) {
    case "paragraph":
      return `<w:p>${inlineXml(block.content, out)}</w:p>`;
    case "heading": {
      const level = Math.min(6, Math.max(1, Math.trunc(block.level) || 1));
      return `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>${inlineXml(block.content, out)}</w:p>`;
    }
    case "list":
      return listXml(block, depth, out);
    case "table":
      return tableXml(block, out);
  }
}

/** Write a document as a Word file (.docx). */
export function writeDocx(doc: TextDoc, limits?: Partial<Limits>): WriteResult {
  const caps = withLimits(limits);
  const out: Out = { rels: [], media: [], nums: [], notes: [], warnings: new Set() };
  const body = doc.blocks.map((b) => blockXml(b, out)).join("");
  const document =
    `${XML_HEADER}<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const numbering =
    `${XML_HEADER}<w:numbering xmlns:w="${W}">${abstractNum(0, false)}${abstractNum(1, true)}` +
    out.nums
      .map((ordered, i) => `<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${ordered ? 1 : 0}"/>${Array.from({ length: 9 }, (_, l) => `<w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="1"/></w:lvlOverride>`).join("")}</w:num>`)
      .join("") +
    `</w:numbering>`;
  const separators = '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>';
  const footnotes = `${XML_HEADER}<w:footnotes ${NS}>${separators}${out.notes.join("")}</w:footnotes>`;
  const overrides: Record<string, string> = {
    "word/document.xml": `${WML}.document.main+xml`,
    "word/styles.xml": `${WML}.styles+xml`,
    "word/numbering.xml": `${WML}.numbering+xml`,
    "word/footnotes.xml": `${WML}.footnotes+xml`,
  };
  const rootRels: { id: string; type: string; target: string }[] = [{ id: "rId1", type: "officeDocument", target: "word/document.xml" }];
  const entries: ZipEntry[] = [];
  if (doc.title) {
    overrides["docProps/core.xml"] = "application/vnd.openxmlformats-package.core-properties+xml";
    rootRels.push({ id: "rId2", type: "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties", target: "docProps/core.xml" });
    entries.push({
      name: "docProps/core.xml",
      data: `${XML_HEADER}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${esc(doc.title)}</dc:title></cp:coreProperties>`,
    });
  }
  const bytes = writeZip(
    [
      {
        name: "[Content_Types].xml",
        data: contentTypesXml({ rels: "application/vnd.openxmlformats-package.relationships+xml", xml: "application/xml", png: "image/png", jpg: "image/jpeg", gif: "image/gif" }, overrides),
      },
      { name: "_rels/.rels", data: relsXml(rootRels) },
      { name: "word/document.xml", data: document },
      {
        name: "word/_rels/document.xml.rels",
        data: relsXml([{ id: "rId1", type: "styles", target: "styles.xml" }, { id: "rId2", type: "numbering", target: "numbering.xml" }, { id: "rId3", type: "footnotes", target: "footnotes.xml" }, ...out.rels]),
      },
      { name: "word/styles.xml", data: STYLES },
      { name: "word/numbering.xml", data: numbering },
      { name: "word/footnotes.xml", data: footnotes },
      ...entries,
      ...out.media,
    ],
    caps,
  );
  return { bytes, warnings: [...out.warnings] };
}

// ---------- read ----------

const MONO = /courier|consolas|mono|menlo|lucida console/i;
/** A toggle property such as `<w:b/>` is on unless its value says off. */
const on = (el: XmlElement | undefined): boolean => !!el && !/^(0|false|off|none)$/i.test(el.attrs["w:val"] ?? "");

/** Read a Word file (.docx). Content with no place in the model is listed in `warnings`. */
export function readDocx(bytes: Uint8Array, limits?: Partial<Limits>): ReadResult<TextDoc> {
  const caps = withLimits(limits);
  const pkg = readOoxmlPackage(bytes, caps, "DOCX", WML);
  const warnings = new Set<string>();
  const warn = (what: string) => warnings.add(`${what}: not imported`);
  const document = pkg.parse(pkg.mainPart);
  const body = childOf(document, "w:body");
  if (!body) throw new FormatError("INVALID_FILE", "the DOCX file has no body");
  const rels = pkg.rels(pkg.mainPart);
  const partOf = (kind: string) => [...rels.values()].find((r) => r.kind === kind && !r.external)?.target;
  const title = (() => {
    const core = [...pkg.rels("").values()].find((r) => r.kind === "core-properties" && !r.external)?.target;
    return core && pkg.files[core] ? textOf(childOf(pkg.parse(core), "dc:title")).trim() : "";
  })();

  const styleNames = new Map<string, string>();
  const stylesPart = partOf("styles");
  if (stylesPart && pkg.files[stylesPart]) for (const s of elementsOf(pkg.parse(stylesPart))) if (s.name === "w:style") styleNames.set(s.attrs["w:styleId"], childOf(s, "w:name")?.attrs["w:val"] ?? "");
  const headingLevel = (styleId: string | undefined): Heading["level"] | undefined => {
    const name = styleNames.get(styleId ?? "") || styleId || "";
    const found = /^heading ?([1-6])$/i.exec(name) ?? /^heading ?([1-6])$/i.exec(styleId ?? "");
    return found ? (Number(found[1]) as Heading["level"]) : /^title$/i.test(name) ? 1 : undefined;
  };

  const abstractLevels = new Map<string, Map<string, string>>();
  const numToAbstract = new Map<string, string>();
  /** A list instance that restarts its numbering at a level starts a new list there. */
  const restarts = new Map<string, Set<string>>();
  const numberingPart = partOf("numbering");
  if (numberingPart && pkg.files[numberingPart]) {
    for (const n of elementsOf(pkg.parse(numberingPart))) {
      if (n.name === "w:abstractNum") abstractLevels.set(n.attrs["w:abstractNumId"], new Map(elementsOf(n).filter((l) => l.name === "w:lvl").map((l) => [l.attrs["w:ilvl"], childOf(l, "w:numFmt")?.attrs["w:val"] ?? "bullet"])));
      else if (n.name === "w:num") {
        numToAbstract.set(n.attrs["w:numId"], childOf(n, "w:abstractNumId")?.attrs["w:val"] ?? "");
        restarts.set(n.attrs["w:numId"], new Set(elementsOf(n).filter((o) => o.name === "w:lvlOverride" && childOf(o, "w:startOverride")).map((o) => o.attrs["w:ilvl"])));
      }
    }
  }
  const isOrdered = (numId: string, ilvl: string): boolean => {
    const fmt = abstractLevels.get(numToAbstract.get(numId) ?? "")?.get(ilvl) ?? "bullet";
    return fmt !== "bullet" && fmt !== "none";
  };

  const marksOf = (rPr: XmlElement | undefined): Mark[] => {
    if (!rPr) return [];
    const marks: Mark[] = [];
    const rStyle = childOf(rPr, "w:rStyle")?.attrs["w:val"];
    if (rStyle === "Strong") marks.push("bold");
    if (rStyle === "Emphasis") marks.push("italic");
    if (on(childOf(rPr, "w:b"))) marks.push("bold");
    if (on(childOf(rPr, "w:i"))) marks.push("italic");
    if (on(childOf(rPr, "w:u"))) marks.push("underline");
    if (on(childOf(rPr, "w:strike")) || on(childOf(rPr, "w:dstrike"))) marks.push("strike");
    if (MONO.test(childOf(rPr, "w:rFonts")?.attrs["w:ascii"] ?? "")) marks.push("code");
    const align = childOf(rPr, "w:vertAlign")?.attrs["w:val"];
    if (align === "superscript") marks.push("superscript");
    else if (align === "subscript") marks.push("subscript");
    return MARKS.filter((m) => marks.includes(m));
  };

  function drawing(el: XmlElement): Inline[] {
    const find = (node: XmlElement, name: string): XmlElement | undefined => (node.name === name ? node : elementsOf(node).map((k) => find(k, name)).find(Boolean));
    const blip = find(el, "a:blip");
    if (!blip) return warn("a drawing that is not a picture"), [];
    if (blip.attrs["r:link"]) return warnings.add("a remote image: not loaded"), [];
    const target = rels.get(blip.attrs["r:embed"] ?? "");
    const data = target && !target.external ? pkg.files[target.target] : undefined;
    const sniffed = data ? sniffImage(data) : undefined;
    if (!data || !sniffed) return warnings.add("an image that is not a PNG, JPEG, or GIF: not imported"), [];
    const extent = find(el, "wp:extent");
    const [w, h] = [Number(extent?.attrs.cx), Number(extent?.attrs.cy)].map((n) => (Number.isFinite(n) && n > 0 ? Math.round((n / EMU) * 10) / 10 : undefined));
    const docPr = find(el, "wp:docPr");
    const alt = (docPr?.attrs.descr || docPr?.attrs.title || "").trim();
    return [{ type: "image", data, mime: sniffed.mime, ...(alt ? { alt } : {}), ...(w && h ? { width: w, height: h } : {}) }];
  }

  const notes = new Map<string, Footnote>();
  const notesPart = partOf("footnotes");
  const readNote = (note: XmlElement): Footnote => {
    const content = elementsOf(note)
      .filter((p) => p.name === "w:p")
      .flatMap((p, i): Inline[] => [...(i ? [{ type: "break" } as const] : []), ...inline(p, [], undefined)]);
    return { type: "footnote", content: content.filter((n): n is TextRun | { type: "break" } => n.type === "text" || n.type === "break") };
  };

  function inline(parent: XmlElement, marks: Mark[], href: string | undefined): Inline[] {
    return elementsOf(parent).flatMap((c): Inline[] => {
      switch (c.name) {
        case "w:r": {
          const runMarks = [...new Set([...marks, ...marksOf(childOf(c, "w:rPr"))])];
          return elementsOf(c).flatMap((k): Inline[] => {
            // A viewer underlines a link by itself, so a link run carries no underline mark.
            const shown = href ? runMarks.filter((m) => m !== "underline") : runMarks;
            const text = (value: string): Inline[] => (value ? [{ type: "text", text: value, ...(shown.length ? { marks: shown } : {}), ...(href ? { href } : {}) }] : []);
            switch (k.name) {
              case "w:t":
                return text(textOf(k));
              case "w:tab":
                return text("\t");
              case "w:noBreakHyphen":
                return text("-");
              case "w:br":
              case "w:cr":
                return k.attrs["w:type"] === "page" || k.attrs["w:type"] === "column" ? [] : [{ type: "break" }];
              case "w:drawing":
                return drawing(k);
              case "w:footnoteReference": {
                const note = notes.get(k.attrs["w:id"]);
                return note ? [note] : (warn("a footnote that is missing"), []);
              }
              case "w:endnoteReference":
                return warn("an endnote"), [];
              case "w:pict":
              case "w:object":
                return warn(k.name), [];
              default:
                return [];
            }
          });
        }
        case "w:hyperlink": {
          const target = c.attrs["w:anchor"] ? `#${c.attrs["w:anchor"]}` : rels.get(c.attrs["r:id"] ?? "")?.target;
          const safe = safeHref(target);
          if (target !== undefined && !safe) warnings.add("a link with an unsafe target was read as plain text");
          return inline(c, marks, safe ?? href);
        }
        case "w:ins":
        case "w:smartTag":
        case "w:sdt":
        case "w:sdtContent":
        case "w:fldSimple":
          return inline(c, marks, href);
        case "w:del":
          return warnings.add("tracked deletions: skipped"), [];
        case "mc:AlternateContent":
          return warn(c.name), [];
        default:
          return [];
      }
    });
  }
  const inlineOf = (p: XmlElement): Inline[] => normalizeInline(inline(p, [], undefined));
  if (notesPart && pkg.files[notesPart]) {
    for (const n of elementsOf(pkg.parse(notesPart))) if (n.name === "w:footnote" && !n.attrs["w:type"]) notes.set(n.attrs["w:id"], readNote(n));
  }

  function blocksOf(parent: XmlElement): Block[] {
    const out: Block[] = [];
    const stack: { list: List; level: number; numId: string }[] = [];
    const item = (p: XmlElement, numId: string, ilvl: string): void => {
      const level = Number(ilvl) || 0;
      const ordered = isOrdered(numId, ilvl);
      while (stack.length && stack[stack.length - 1].level > level) stack.pop();
      let top = stack[stack.length - 1];
      if (top && level === 0 && top.numId !== numId && restarts.get(numId)?.has("0")) top = undefined as never;
      if (!top || top.level < level) {
        const list: List = { type: "list", ordered, items: [] };
        if (top) {
          const parentItem = top.list.items[top.list.items.length - 1];
          parentItem.sublist = list;
        } else {
          stack.length = 0;
          out.push(list);
        }
        stack.push((top = { list, level: top ? top.level + 1 : level, numId }));
      } else if (top.list.ordered !== ordered) warnings.add("a list that mixes ordered and plain items kept the type of its first item");
      top.list.items.push({ content: inlineOf(p) });
    };
    for (const el of elementsOf(parent)) {
      if (el.name === "w:p") {
        const pPr = childOf(el, "w:pPr");
        const level = headingLevel(childOf(pPr, "w:pStyle")?.attrs["w:val"]);
        const numPr = childOf(pPr, "w:numPr");
        const numId = childOf(numPr, "w:numId")?.attrs["w:val"];
        if (level === undefined && numId && numId !== "0") item(el, numId, childOf(numPr, "w:ilvl")?.attrs["w:val"] ?? "0");
        else {
          stack.length = 0;
          out.push(level ? { type: "heading", level, content: inlineOf(el) } : { type: "paragraph", content: inlineOf(el) });
        }
      } else if (el.name === "w:tbl") {
        stack.length = 0;
        const rows = elementsOf(el).filter((r) => r.name === "w:tr");
        if (elementsOf(el).some((r) => r.name === "w:tr" && elementsOf(r).some((c) => c.name === "w:tc" && (childOf(childOf(c, "w:tcPr"), "w:gridSpan") || childOf(childOf(c, "w:tcPr"), "w:vMerge"))))) warnings.add("merged table cells were read as single cells");
        if (rows.length) {
          const header = !!childOf(childOf(rows[0], "w:trPr"), "w:tblHeader");
          out.push({ type: "table", ...(header ? { header: true } : {}), rows: rows.map((r) => elementsOf(r).filter((c) => c.name === "w:tc").map((c) => blocksOf(c))) });
        }
      } else if (el.name === "w:sdt") out.push(...blocksOf(childOf(el, "w:sdtContent") ?? el));
      else if (el.name === "w:sdtContent") out.push(...blocksOf(el));
    }
    return out;
  }

  const doc: TextDoc = { ...(title ? { title } : {}), blocks: blocksOf(body) };
  return { model: normalizeTextDoc(doc), warnings: [...warnings] };
}
