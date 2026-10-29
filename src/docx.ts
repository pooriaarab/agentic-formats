import { frameSize, sniffImage } from "./image.js";
import { FormatError, capped, withLimits, type Limits } from "./limits.js";
import {
  MARKS,
  checkList,
  checkObject,
  checkRun,
  headingLevel,
  kindName,
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
  /** Hyperlinks inside footnotes: relationships of the footnotes part, which has ids of its own. */
  noteRels: { id: string; type: string; target: string; external?: boolean }[];
  inNote: boolean;
  /** Lists and tables that are open now, and the most the writer allows. A deep or cyclic model fails before the stack does. */
  depth: number;
  maxDepth: number;
  media: ZipEntry[];
  nums: boolean[]; // one entry per list; true when ordered
  notes: string[];
  warnings: Set<string>;
  /** Counts the text of each run and the length of each link as it is written, so a model that holds one note many times stops at the cap. */
  charge: (chars: number) => void;
  /** Counts the XML as each piece is made, so a model that holds one object many times fails before one string grows past the cap. */
  bytes: (length: number) => void;
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

/** The most characters one run can need, found without making them: its text, the tags around each piece, and a tag for each tab or newline. */
function runBound(text: string): number {
  let n = text.length + 160;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "\t" || c === "\n") n += 48;
    else if (c === "&" || c === "<" || c === ">" || c === '"' || c === "\r") n += 5;
  }
  return n;
}

function runXml(run: TextRun, out: Out): string {
  checkRun(run);
  out.charge(run.text.length + (run.href?.length ?? 0));
  // The worst case counts before the string exists: a tab or a newline grows to a tag, a quote to six characters, and a link adds a relationship.
  out.bytes(runBound(run.text) + 6 * (run.href?.length ?? 0) + (run.href === undefined ? 0 : 200));
  const href = run.href === undefined ? undefined : safeHref(run.href);
  if (run.href !== undefined && !href) out.warnings.add("a link with an unsafe target was written as plain text");
  const props = (href ? '<w:rStyle w:val="Hyperlink"/>' : "") + RPR_ORDER.filter((m) => run.marks?.includes(m)).map((m) => MARK_XML[m]).join("");
  const xml = `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ""}${textPieces(run.text)}</w:r>`;
  if (!href) return xml;
  if (href.startsWith("#")) return `<w:hyperlink w:anchor="${esc(href.slice(1))}">${xml}</w:hyperlink>`;
  // A footnote link belongs to the footnotes part, whose relationships are not the document's.
  const rels = out.inNote ? out.noteRels : out.rels;
  const id = out.inNote ? `rId${rels.length + 1}` : nextId(out);
  rels.push({ id, type: "hyperlink", target: href, external: true });
  return `<w:hyperlink r:id="${id}">${xml}</w:hyperlink>`;
}

/** The runs of a footnote. Their links go to the footnotes part. */
function noteContent(nodes: Inline[], out: Out): string {
  const before = out.inNote;
  out.inNote = true;
  try {
    return inlineXml(nodes, out);
  } finally {
    out.inNote = before;
  }
}

function inlineXml(nodes: Inline[], out: Out): string {
  if (!Array.isArray(nodes)) throw new FormatError("INVALID_MODEL", "the content of a block must be a list");
  return nodes
    .map((node) => {
      const xml = ((): string => {
        if (typeof node !== "object" || node === null) throw new FormatError("INVALID_MODEL", "a piece of inline content is not an object");
        switch (node.type) {
          case "text":
            return runXml(node, out);
          case "break":
            return "<w:r><w:br/></w:r>";
          case "image":
            return imageXml(node, out);
          case "footnote": {
            // A note holds text and line breaks only. A note inside a note would get the same id as its parent, and a cyclic one would recurse forever.
            if (!Array.isArray(node.content) || node.content.some((inner) => typeof inner !== "object" || inner === null || (inner.type !== "text" && inner.type !== "break"))) {
              throw new FormatError("INVALID_MODEL", "a footnote holds only text and line breaks");
            }
            const id = out.notes.length + 1;
            const note = `<w:footnote w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> </w:t></w:r>${noteContent(node.content, out)}</w:p></w:footnote>`;
            out.bytes(note.length);
            out.notes.push(note);
            return `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>`;
          }
          default:
            throw new FormatError("INVALID_MODEL", `${kindName((node as { type?: unknown }).type)} is not a kind of inline content`);
        }
      })();
      // A run charged its own XML already.
      if (node.type !== "text") out.bytes(xml.length);
      return xml;
    })
    .join("");
}

/** One list instance in numbering.xml. Each instance restarts its numbering at every level. */
const numXml = (id: number, ordered: boolean): string =>
  `<w:num w:numId="${id}"><w:abstractNumId w:val="${ordered ? 1 : 0}"/>${Array.from({ length: 9 }, (_, l) => `<w:lvlOverride w:ilvl="${l}"><w:startOverride w:val="1"/></w:lvlOverride>`).join("")}</w:num>`;

/** Run `body` one level deeper. */
function nested<T>(out: Out, body: () => T): T {
  if (out.depth >= out.maxDepth) throw new FormatError("LIMIT_EXCEEDED", `the document nests lists and tables deeper than ${out.maxDepth} levels`);
  out.depth++;
  try {
    return body();
  } finally {
    out.depth--;
  }
}

function listXml(list: List, depth: number, out: Out): string {
  return nested(out, () => listXmlOf(list, depth, out));
}

function listXmlOf(list: List, depth: number, out: Out): string {
  checkList(list.items, "the items of a list");
  for (const item of list.items) checkObject(item, "an item of a list");
  if (depth > 8) out.warnings.add("a list nests deeper than nine levels: the deepest levels were flattened");
  const level = Math.min(depth, 8);
  const numId = out.nums.push(list.ordered);
  // Every list adds an entry to numbering.xml, even an empty one.
  out.bytes(numXml(numId, list.ordered).length);
  return list.items
    .map((item) => {
      const xml =
        `<w:p><w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="${level}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>${inlineXml(item.content, out)}</w:p>` +
        (item.sublist ? listXml(item.sublist, depth + 1, out) : "");
      out.bytes(xml.length);
      return xml;
    })
    .join("");
}

function tableXml(table: Table, out: Out): string {
  return nested(out, () => tableXmlOf(table, out));
}

function tableXmlOf(table: Table, out: Out): string {
  checkList(table.rows, "the rows of a table");
  for (const row of table.rows) {
    checkList(row, "a row of a table");
    for (const cell of row) checkList(cell, "a cell of a table");
  }
  const columns = table.rows.reduce((n, r) => Math.max(n, r.length), 1);
  const width = Math.floor(9000 / columns);
  const row = (cells: Block[][], header: boolean) =>
    `<w:tr>${header ? "<w:trPr><w:tblHeader/></w:trPr>" : ""}${Array.from({ length: columns }, (_, i) => {
      const blocks = (cells[i] ?? []).map((b) => blockXml(b, out)).join("");
      const last = cells[i]?.[cells[i].length - 1];
      const tc = `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${blocks}${!last || last.type === "table" ? "<w:p/>" : ""}</w:tc>`;
      out.bytes(tc.length);
      return tc;
    }).join("")}</w:tr>`;
  return (
    `<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid>${`<w:gridCol w:w="${width}"/>`.repeat(columns)}</w:tblGrid>` +
    `${table.rows.map((cells, i) => row(cells, !!table.header && i === 0)).join("")}</w:tbl>`
  );
}

function blockXml(block: Block, out: Out, depth = 0): string {
  const xml = blockXmlOf(block, out, depth);
  out.bytes(xml.length);
  return xml;
}

function blockXmlOf(block: Block, out: Out, depth: number): string {
  checkObject(block, "a block");
  switch (block.type) {
    case "paragraph":
      return `<w:p>${inlineXml(block.content, out)}</w:p>`;
    case "heading": {
      const level = headingLevel(block.level);
      return `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>${inlineXml(block.content, out)}</w:p>`;
    }
    case "list":
      return listXml(block, depth, out);
    case "table":
      return tableXml(block, out);
    default:
      throw new FormatError("INVALID_MODEL", `${kindName((block as { type?: unknown } | null)?.type)} is not a kind of block`);
  }
}

/** Write a document as a Word file (.docx). */
export function writeDocx(doc: TextDoc, limits?: Partial<Limits>): WriteResult {
  const caps = withLimits(limits);
  const out: Out = {
    rels: [],
    noteRels: [],
    inNote: false,
    depth: 0,
    maxDepth: caps.maxXmlDepth,
    media: [],
    nums: [],
    notes: [],
    warnings: new Set(),
    charge: capped(caps.maxTextChars, `the document has more than ${caps.maxTextChars} characters of text`),
    bytes: capped(caps.maxTotalBytes, `the document passes ${caps.maxTotalBytes} bytes while it is written`),
  };
  checkList(doc?.blocks, "the blocks of a document");
  const body = doc.blocks.map((b) => blockXml(b, out)).join("");
  const document =
    `${XML_HEADER}<w:document ${NS}><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`;
  const numbering =
    `${XML_HEADER}<w:numbering xmlns:w="${W}">${abstractNum(0, false)}${abstractNum(1, true)}` +
    out.nums
      .map((ordered, i) => numXml(i + 1, ordered))
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
      ...(out.noteRels.length ? [{ name: "word/_rels/footnotes.xml.rels", data: relsXml(out.noteRels) }] : []),
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
    const target = activeRels.get(blip.attrs["r:embed"] ?? "");
    const data = target && !target.external ? pkg.files[target.target] : undefined;
    const sniffed = data ? sniffImage(data) : undefined;
    if (!data || !sniffed) return warnings.add("an image that is not a PNG, JPEG, or GIF: not imported"), [];
    chargeImages(data.length);
    const extent = find(el, "wp:extent");
    const [w, h] = [Number(extent?.attrs.cx), Number(extent?.attrs.cy)].map((n) => (Number.isFinite(n) && n > 0 ? Math.round((n / EMU) * 10) / 10 : undefined));
    const docPr = find(el, "wp:docPr");
    const alt = (docPr?.attrs.descr || docPr?.attrs.title || "").trim();
    return [{ type: "image", data, mime: sniffed.mime, ...(alt ? { alt } : {}), ...(w && h ? { width: w, height: h } : {}) }];
  }

  const notes = new Map<string, Footnote>();
  /** An image part is shared by every drawing that names it, so each drawing counts its bytes. */
  const chargeRuns = capped(caps.maxTextChars, `the file has more than ${caps.maxTextChars} characters of run text and link targets`);
  const chargeImages = capped(caps.maxTotalBytes, `the images in the file add up to more than ${caps.maxTotalBytes} bytes`);
  /** Notes already placed. A note belongs to one reference; a repeat would copy its whole content again. */
  const placed = new Set<string>();
  const notesPart = partOf("footnotes");
  /** The relationships that links in the part being read resolve against: the document's, or the footnotes part's. */
  let activeRels = rels;
  // Read once: every note resolves its links against the same part.
  const noteRelations = notesPart ? pkg.rels(notesPart) : rels;
  const readNote = (note: XmlElement): Footnote => {
    const before = activeRels;
    activeRels = noteRelations;
    try {
      const content = elementsOf(note)
        .filter((p) => p.name === "w:p")
        .flatMap((p, i): Inline[] => [...(i ? [{ type: "break" } as const] : []), ...inline(p, [], undefined)]);
      const kept = content.filter((n): n is TextRun | { type: "break" } => n.type === "text" || n.type === "break");
      if (kept.length !== content.length) warn("an image or a note inside a footnote");
      return { type: "footnote", content: kept };
    } finally {
      activeRels = before;
    }
  };

  function inline(parent: XmlElement, marks: Mark[], href: string | undefined): Inline[] {
    return elementsOf(parent).flatMap((c): Inline[] => {
      switch (c.name) {
        case "w:r": {
          const runMarks = [...new Set([...marks, ...marksOf(childOf(c, "w:rPr"))])];
          return elementsOf(c).flatMap((k): Inline[] => {
            // A viewer underlines a link by itself, so a link run carries no underline mark.
            const shown = href ? runMarks.filter((m) => m !== "underline") : runMarks;
            const text = (value: string): Inline[] => {
              if (!value) return [];
              // A link target is one string in a shared part, but each run that names it carries it, so each run counts it.
              chargeRuns(value.length + (href?.length ?? 0));
              return [{ type: "text", text: value, ...(shown.length ? { marks: shown } : {}), ...(href ? { href } : {}) }];
            };
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
                const id = k.attrs["w:id"];
                const note = notes.get(id);
                if (!note) return warn("a footnote that is missing"), [];
                if (placed.has(id)) return warn("a repeated footnote reference"), [];
                placed.add(id);
                return [note];
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
          const rel = c.attrs["r:id"] ? activeRels.get(c.attrs["r:id"]) : undefined;
          // Only an external hyperlink relationship is a link: an id that names the styles part, or a missing id, is not.
          const target = c.attrs["w:anchor"] ? `#${c.attrs["w:anchor"]}` : rel?.kind === "hyperlink" && rel.external ? rel.target : undefined;
          if (c.attrs["r:id"] && !c.attrs["w:anchor"] && target === undefined) warnings.add("a link whose relationship is missing or is not a link was read as plain text");
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
      // A level is one digit from 0 to 8. Anything else reads as level 0, with a warning.
      const valid = /^[0-8]$/.test(ilvl);
      if (!valid) warnings.add("a list level that is not 0 to 8 was read as level 0");
      const level = valid ? Number(ilvl) : 0;
      const ordered = isOrdered(numId, String(level));
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
      } else if (el.name === "w:sdt" || el.name === "w:sdtContent") {
        for (const inner of blocksOf(el.name === "w:sdt" ? (childOf(el, "w:sdtContent") ?? el) : el)) out.push(inner);
      }
    }
    return out;
  }

  const doc: TextDoc = { ...(title ? { title } : {}), blocks: blocksOf(body) };
  return { model: normalizeTextDoc(doc), warnings: [...warnings] };
}
