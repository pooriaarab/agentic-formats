import { MAX_POINTS, checkDeck, tenth, textBound } from "./deck.js";
import { sniffImage } from "./image.js";
import { FormatError, capped, withLimits, type Limits } from "./limits.js";
import { MARKS, normalizeInline, type Deck, type Frame, type Inline, type Mark, type ReadResult, type Slide, type SlideShape, type TextRun, type WriteResult } from "./model.js";
import { REL_NS, XML_HEADER, contentTypesXml, readOoxmlPackage, relsXml } from "./ooxml.js";
import { childOf, elementsOf, safeHref, textOf, xmlEscape as esc, type XmlElement } from "./xml.js";
import { writeZip, type ZipEntry } from "./zip.js";

const A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PML = "application/vnd.openxmlformats-officedocument.presentationml";
const NS = `xmlns:a="${A}" xmlns:r="${REL_NS}" xmlns:p="${P}"`;
const EMU = 12700;
const emu = (pt: number): number => Math.round(pt * EMU);
const MONO = /courier|consolas|mono|menlo|lucida console/i;

// ---------- write ----------

const solid = (color: string) => `<a:solidFill><a:srgbClr val="${color.slice(1).toUpperCase()}"/></a:solidFill>`;
const xfrm = (f: Frame, tag = "a:xfrm") => `<${tag}><a:off x="${emu(f.x)}" y="${emu(f.y)}"/><a:ext cx="${emu(f.width)}" cy="${emu(f.height)}"/></${tag}>`;
const GROUP = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;
const CLR = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
const COLORS = ["000000", "FFFFFF", "44546A", "E7E6E6", "4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5", "70AD47", "0563C1", "954F72"];
const THEME =
  `${XML_HEADER}<a:theme xmlns:a="${A}" name="Office"><a:themeElements><a:clrScheme name="Office">` +
  CLR.map((name, i) => `<a:${name}>${i === 0 ? '<a:sysClr val="windowText" lastClr="000000"/>' : i === 1 ? '<a:sysClr val="window" lastClr="FFFFFF"/>' : `<a:srgbClr val="${COLORS[i]}"/>`}</a:${name}>`).join("") +
  `</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
  `<a:fmtScheme name="Office"><a:fillStyleLst>${'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3)}</a:fillStyleLst>` +
  `<a:lnStyleLst>${'<a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'.repeat(3)}</a:lnStyleLst>` +
  `<a:effectStyleLst>${"<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3)}</a:effectStyleLst>` +
  `<a:bgFillStyleLst>${'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
const CLR_MAP = `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>`;
const MASTER =
  `${XML_HEADER}<p:sldMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GROUP}</p:spTree></p:cSld>${CLR_MAP}` +
  `<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>`;
const LAYOUT = `${XML_HEADER}<p:sldLayout ${NS} type="blank" preserve="1"><p:cSld name="Blank"><p:spTree>${GROUP}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
const placeholder = (id: number, name: string, ph: string, f: Frame, extra = "") =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"${ph.startsWith("sldImg") ? ' noRot="1" noChangeAspect="1"' : ""}/></p:cNvSpPr><p:nvPr><p:ph ${ph}/></p:nvPr></p:nvSpPr><p:spPr>${xfrm(f)}${ph.startsWith("sldImg") ? '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' : ""}</p:spPr>${extra}</p:sp>`;
const NOTES_BODY = { x: 54, y: 346.5, width: 432, height: 283.5 };
const NOTES_IMAGE = { x: 0, y: 0, width: 540, height: 720 };
const NOTES_MASTER =
  `${XML_HEADER}<p:notesMaster ${NS}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GROUP}` +
  placeholder(2, "Slide Image Placeholder 1", 'type="sldImg" idx="2"', { x: 90, y: 54, width: 360, height: 202.5 }) +
  placeholder(3, "Notes Placeholder 2", 'type="body" sz="quarter" idx="3"', NOTES_BODY, '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>') +
  `</p:spTree></p:cSld>${CLR_MAP}</p:notesMaster>`;

interface Out {
  warnings: Set<string>;
}
interface SlideOut extends Out {
  /** Counts the XML as each piece is made, so a model that holds one shape many times fails before one string grows past the cap. */
  bytes: (length: number) => void;
  rels: { id: string; type: string; target: string; external?: boolean }[];
  media: ZipEntry[];
  mediaCount: { n: number };
  nextShape: number;
}

function runXml(run: TextRun, out: SlideOut): string {
  // A newline grows to a tag of 41 characters, and each line gets a run of about 260. The worst case counts before the string exists.
  out.bytes(textBound(run.text, 300) + (run.href === undefined || run.text === "" ? 0 : 200 + 6 * run.href.length));
  const marks = MARKS.filter((m) => run.marks?.includes(m));
  const attrs = `lang="en-US" sz="1800"${marks.includes("bold") ? ' b="1"' : ""}${marks.includes("italic") ? ' i="1"' : ""}${marks.includes("underline") ? ' u="sng"' : ""}${marks.includes("strike") ? ' strike="sngStrike"' : ""}${marks.includes("superscript") ? ' baseline="30000"' : marks.includes("subscript") ? ' baseline="-25000"' : ""}`;
  const href = run.href === undefined ? undefined : safeHref(run.href);
  if (run.href !== undefined && !href) out.warnings.add("a link with an unsafe target was written as plain text");
  let link = "";
  if (href?.startsWith("#") && run.text !== "") out.warnings.add("an in-page link was written as plain text");
  // A run with no text writes nothing, so it gets no relationship either.
  if (href && !href.startsWith("#") && run.text !== "") {
    const id = `rId${out.rels.length + 3}`; // rId1 is the layout, rId2 the notes
    out.rels.push({ id, type: "hyperlink", target: href, external: true });
    link = `<a:hlinkClick r:id="${id}"/>`;
  }
  // Every run names its font: LibreOffice gives a run with no font the font of the run before it.
  const code = marks.includes("code") ? '<a:latin typeface="Courier New"/><a:cs typeface="Courier New"/>' : '<a:latin typeface="Calibri"/><a:cs typeface="Calibri"/>';
  const xml = run.text
    .split(/(\n)/)
    .map((piece) => (piece === "\n" ? "<a:br><a:rPr lang=\"en-US\" sz=\"1800\"/></a:br>" : piece ? `<a:r><a:rPr ${attrs}>${code}${link}</a:rPr><a:t>${esc(piece)}</a:t></a:r>` : ""))
    .join("");
  return xml;
}

function shapeXml(shape: SlideShape, out: SlideOut): string {
  const xml = shapeXmlOf(shape, out);
  out.bytes(xml.length);
  return xml;
}

function shapeXmlOf(shape: SlideShape, out: SlideOut): string {
  // A description is escaped, which grows it up to six times, so it counts before the string exists.
  if (shape.type === "picture") out.bytes(6 * (shape.alt?.length ?? 0));
  const id = ++out.nextShape;
  switch (shape.type) {
    case "textbox":
      return (
        `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(shape.frame)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>${shape.paragraphs
          .map((runs) => {
            const p = `<a:p>${runs.map((r) => runXml(r, out)).join("") || '<a:endParaRPr lang="en-US" sz="1800"/>'}</a:p>`;
            out.bytes(p.length);
            return p;
          })
          .join("")}</p:txBody></p:sp>`
      );
    case "picture": {
      const sniffed = sniffImage(shape.data);
      if (!sniffed) {
        out.warnings.add("an image is not a PNG, JPEG, or GIF: not written");
        return "";
      }
      const n = ++out.mediaCount.n;
      const rel = `rId${out.rels.length + 3}`;
      out.media.push({ name: `ppt/media/image${n}.${sniffed.ext}`, data: shape.data });
      out.rels.push({ id: rel, type: "image", target: `../media/image${n}.${sniffed.ext}` });
      return (
        `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}" descr="${esc(shape.alt ?? "")}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${rel}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(shape.frame)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
      );
    }
    case "shape": {
      const line = shape.stroke ? `<a:ln w="12700">${solid(shape.stroke)}</a:ln>` : "<a:ln><a:noFill/></a:ln>";
      if (shape.kind === "line") {
        return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Line ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr>${xfrm(shape.frame)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${line}</p:spPr></p:cxnSp>`;
      }
      const fill = shape.fill ? solid(shape.fill) : "<a:noFill/>";
      return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(shape.frame)}<a:prstGeom prst="${shape.kind === "ellipse" ? "ellipse" : "rect"}"><a:avLst/></a:prstGeom>${fill}${line}</p:spPr></p:sp>`;
    }
  }
}

/** Write a deck as a PowerPoint file (.pptx). */
export function writePptx(deck: Deck, limits?: Partial<Limits>): WriteResult {
  const caps = withLimits(limits);
  checkDeck(deck, caps);
  const warnings = new Set<string>();
  const media: ZipEntry[] = [];
  const mediaCount = { n: 0 };
  const chargeBytes = capped(caps.maxTotalBytes, `the deck passes ${caps.maxTotalBytes} bytes while it is written`);
  const hasNotes = deck.slides.some((s) => s.notes);
  const overrides: Record<string, string> = {
    "ppt/presentation.xml": `${PML}.presentation.main+xml`,
    "ppt/slideMasters/slideMaster1.xml": `${PML}.slideMaster+xml`,
    "ppt/slideLayouts/slideLayout1.xml": `${PML}.slideLayout+xml`,
    "ppt/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
  };
  const entries: ZipEntry[] = [];
  deck.slides.forEach((slide, i) => {
    const n = i + 1;
    const out: SlideOut = { warnings, rels: [], media, mediaCount, nextShape: 1, bytes: chargeBytes };
    const shapes = slide.shapes.map((s) => shapeXml(s, out)).join("");
    entries.push({ name: `ppt/slides/slide${n}.xml`, data: `${XML_HEADER}<p:sld ${NS}><p:cSld><p:spTree>${GROUP}${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>` });
    const rels = [{ id: "rId1", type: "slideLayout", target: "../slideLayouts/slideLayout1.xml" }, ...(slide.notes ? [{ id: "rId2", type: "notesSlide", target: `../notesSlides/notesSlide${n}.xml` }] : []), ...out.rels];
    entries.push({ name: `ppt/slides/_rels/slide${n}.xml.rels`, data: relsXml(rels) });
    overrides[`ppt/slides/slide${n}.xml`] = `${PML}.slide+xml`;
    if (slide.notes) {
      chargeBytes(textBound(slide.notes, 80));
      const paragraphs = slide.notes.split("\n").map((line) => `<a:p>${line ? `<a:r><a:rPr lang="en-US"/><a:t>${esc(line)}</a:t></a:r>` : '<a:endParaRPr lang="en-US"/>'}</a:p>`).join("");
      entries.push({
        name: `ppt/notesSlides/notesSlide${n}.xml`,
        data:
          `${XML_HEADER}<p:notes ${NS}><p:cSld><p:spTree>${GROUP}` +
          placeholder(2, "Slide Image Placeholder 1", 'type="sldImg"', NOTES_IMAGE) +
          placeholder(3, "Notes Placeholder 2", 'type="body" idx="1"', NOTES_BODY, `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody>`) +
          `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`,
      });
      entries.push({ name: `ppt/notesSlides/_rels/notesSlide${n}.xml.rels`, data: relsXml([{ id: "rId1", type: "notesMaster", target: "../notesMasters/notesMaster1.xml" }, { id: "rId2", type: "slide", target: `../slides/slide${n}.xml` }]) });
      overrides[`ppt/notesSlides/notesSlide${n}.xml`] = `${PML}.notesSlide+xml`;
    }
  });
  const count = deck.slides.length;
  const rels = [
    { id: "rId1", type: "slideMaster", target: "slideMasters/slideMaster1.xml" },
    ...deck.slides.map((_, i) => ({ id: `rId${i + 2}`, type: "slide", target: `slides/slide${i + 1}.xml` })),
    { id: `rId${count + 2}`, type: "theme", target: "theme/theme1.xml" },
    ...(hasNotes ? [{ id: `rId${count + 3}`, type: "notesMaster", target: "notesMasters/notesMaster1.xml" }] : []),
  ];
  if (hasNotes) {
    overrides["ppt/notesMasters/notesMaster1.xml"] = `${PML}.notesMaster+xml`;
    overrides["ppt/theme/theme2.xml"] = "application/vnd.openxmlformats-officedocument.theme+xml";
    entries.push({ name: "ppt/notesMasters/notesMaster1.xml", data: NOTES_MASTER }, { name: "ppt/notesMasters/_rels/notesMaster1.xml.rels", data: relsXml([{ id: "rId1", type: "theme", target: "../theme/theme2.xml" }]) }, { name: "ppt/theme/theme2.xml", data: THEME });
  }
  const presentation =
    `${XML_HEADER}<p:presentation ${NS}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>` +
    (hasNotes ? `<p:notesMasterIdLst><p:notesMasterId r:id="rId${count + 3}"/></p:notesMasterIdLst>` : "") +
    `<p:sldIdLst>${deck.slides.map((_, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join("")}</p:sldIdLst><p:sldSz cx="${emu(deck.width)}" cy="${emu(deck.height)}"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`;
  const bytes = writeZip(
    [
      { name: "[Content_Types].xml", data: contentTypesXml({ rels: "application/vnd.openxmlformats-package.relationships+xml", xml: "application/xml", png: "image/png", jpg: "image/jpeg", gif: "image/gif" }, overrides) },
      { name: "_rels/.rels", data: relsXml([{ id: "rId1", type: "officeDocument", target: "ppt/presentation.xml" }]) },
      { name: "ppt/presentation.xml", data: presentation },
      { name: "ppt/_rels/presentation.xml.rels", data: relsXml(rels) },
      { name: "ppt/slideMasters/slideMaster1.xml", data: MASTER },
      { name: "ppt/slideMasters/_rels/slideMaster1.xml.rels", data: relsXml([{ id: "rId1", type: "slideLayout", target: "../slideLayouts/slideLayout1.xml" }, { id: "rId2", type: "theme", target: "../theme/theme1.xml" }]) },
      { name: "ppt/slideLayouts/slideLayout1.xml", data: LAYOUT },
      { name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels", data: relsXml([{ id: "rId1", type: "slideMaster", target: "../slideMasters/slideMaster1.xml" }]) },
      { name: "ppt/theme/theme1.xml", data: THEME },
      ...entries,
      ...media,
    ],
    caps,
  );
  return { bytes, warnings: [...warnings] };
}

// ---------- read ----------

const hex = (el: XmlElement | undefined): string | undefined => {
  const value = childOf(childOf(el, "a:solidFill"), "a:srgbClr")?.attrs.val;
  return value && /^[0-9a-f]{6}$/i.test(value) ? `#${value.toLowerCase()}` : undefined;
};

/** Read a PowerPoint file (.pptx). A placeholder takes its position from its layout, then its master. */
export function readPptx(bytes: Uint8Array, limits?: Partial<Limits>): ReadResult<Deck> {
  const caps = withLimits(limits);
  const pkg = readOoxmlPackage(bytes, caps, "PPTX", PML);
  const warnings = new Set<string>();
  const warn = (what: string) => warnings.add(`${what}: not imported`);
  const presentation = pkg.parse(pkg.mainPart);
  const rels = pkg.rels(pkg.mainPart);
  const size = childOf(presentation, "p:sldSz");
  const width = Number(size?.attrs.cx) / EMU;
  const height = Number(size?.attrs.cy) / EMU;
  if (!(width > 0 && height > 0 && width <= MAX_POINTS && height <= MAX_POINTS)) throw new FormatError("INVALID_FILE", "the PPTX file has no slide size");

  const kindTarget = (partRels: ReturnType<typeof pkg.rels>, kind: string) => [...partRels.values()].find((r) => r.kind === kind && !r.external)?.target;
  const relTarget = (part: string, kind: string) => kindTarget(pkg.rels(part), kind);
  const treeOf = (part: string | undefined) => (part && pkg.files[part] ? childOf(childOf(pkg.parse(part), "p:cSld"), "p:spTree") : undefined);
  const phKey = (sp: XmlElement) => {
    const ph = childOf(childOf(childOf(sp, "p:nvSpPr"), "p:nvPr"), "p:ph");
    return ph ? `${ph.attrs.type ?? "body"}|${ph.attrs.idx ?? "0"}` : undefined;
  };
  const frameFrom = (xf: XmlElement | undefined): Frame | undefined => {
    const off = childOf(xf, "a:off");
    const ext = childOf(xf, "a:ext");
    if (!off || !ext) return undefined;
    const frame = { x: tenth(Number(off.attrs.x) / EMU), y: tenth(Number(off.attrs.y) / EMU), width: tenth(Number(ext.attrs.cx) / EMU), height: tenth(Number(ext.attrs.cy) / EMU) };
    // A position that is not a number, or a size that is negative or past the model's bound, is not read.
    const usable = [frame.x, frame.y].every((n) => Number.isFinite(n) && Math.abs(n) <= MAX_POINTS) && [frame.width, frame.height].every((n) => Number.isFinite(n) && n >= 0 && n <= MAX_POINTS);
    return usable ? frame : undefined;
  };

  /** Where each placeholder of a layout or master sits, by placeholder key; the first one with a position wins. A part is parsed once, however many slides use it. */
  const placeholderCache = new Map<string, Map<string, Frame>>();
  const placeholdersOf = (part: string | undefined): Map<string, Frame> => {
    const cached = part ? placeholderCache.get(part) : undefined;
    if (cached) return cached;
    const index = new Map<string, Frame>();
    for (const sp of elementsOf(treeOf(part))) {
      const key = sp.name === "p:sp" ? phKey(sp) : undefined;
      const frame = key === undefined ? undefined : frameFrom(childOf(childOf(sp, "p:spPr"), "a:xfrm"));
      if (key !== undefined && frame && !index.has(key)) index.set(key, frame);
    }
    if (part) placeholderCache.set(part, index);
    return index;
  };
  /** A layout, then its master: the order a placeholder looks for its position. */
  const parentsCache = new Map<string, Map<string, Frame>[]>();
  const parentsOf = (layout: string | undefined): Map<string, Frame>[] => {
    const cached = parentsCache.get(layout ?? "");
    if (cached) return cached;
    const parents = [placeholdersOf(layout), placeholdersOf(layout ? relTarget(layout, "slideMaster") : undefined)];
    parentsCache.set(layout ?? "", parents);
    return parents;
  };

  /** Slide and notes parts read so far. A part that two entries name would be read and copied twice. */
  const readParts = new Set<string>();
  // An image part is shared by every picture that names it, so each picture counts its bytes.
  const chargeImages = capped(caps.maxTotalBytes, `the images in the file add up to more than ${caps.maxTotalBytes} bytes`);
  const slides: Slide[] = [];
  const slideEntries = elementsOf(childOf(presentation, "p:sldIdLst"));
  // The writers refuse more slides than this, so a file that lists more is not read.
  if (slideEntries.length > caps.maxEntries / 4) throw new FormatError("LIMIT_EXCEEDED", `the PPTX file lists more than ${caps.maxEntries / 4} slides`);
  for (const entry of slideEntries) {
    const slideRel = rels.get(entry.attrs["r:id"] ?? "");
    if (!slideRel || slideRel.external || !pkg.files[slideRel.target]) {
      warn("a slide that is missing");
      continue;
    }
    const part = slideRel.target;
    if (readParts.has(part)) {
      warn("a slide that is listed twice");
      continue;
    }
    readParts.add(part);
    const slideRels = pkg.rels(part);
    const parents = parentsOf(kindTarget(slideRels, "slideLayout"));
    const inherited = (sp: XmlElement): Frame | undefined => {
      const key = phKey(sp);
      return key === undefined ? undefined : parents.map((index) => index.get(key)).find((frame) => frame !== undefined);
    };

    const paragraph = (p: XmlElement): TextRun[] => {
      const inlines = elementsOf(p).flatMap((r): Inline[] => {
        if (r.name === "a:br") return [{ type: "text", text: "\n" }];
        if (r.name !== "a:r" && r.name !== "a:fld") return [];
        const pr = childOf(r, "a:rPr");
        const marks: Mark[] = [];
        if (pr?.attrs.b === "1" || pr?.attrs.b === "true") marks.push("bold");
        if (pr?.attrs.i === "1" || pr?.attrs.i === "true") marks.push("italic");
        if (pr?.attrs.u && pr.attrs.u !== "none") marks.push("underline");
        if (pr?.attrs.strike && pr.attrs.strike !== "noStrike") marks.push("strike");
        if (MONO.test(childOf(pr, "a:latin")?.attrs.typeface ?? "")) marks.push("code");
        const shift = Number(pr?.attrs.baseline ?? 0);
        if (shift > 0) marks.push("superscript");
        else if (shift < 0) marks.push("subscript");
        const link = childOf(pr, "a:hlinkClick");
        const target = link?.attrs["r:id"] ? slideRels.get(link.attrs["r:id"])?.target : undefined;
        const href = safeHref(target);
        if (target !== undefined && !href) warnings.add("a link with an unsafe target was read as plain text");
        const text = textOf(childOf(r, "a:t"));
        // A viewer underlines a link by itself, so a link run carries no underline mark.
        const shown = MARKS.filter((m) => marks.includes(m) && !(href && m === "underline"));
        return text ? [{ type: "text", text, ...(shown.length ? { marks: shown } : {}), ...(href ? { href } : {}) }] : [];
      });
      return normalizeInline(inlines) as TextRun[];
    };
    const paragraphsOf = (sp: XmlElement) => elementsOf(childOf(sp, "p:txBody")).filter((p) => p.name === "a:p").map(paragraph);

    const shapes: SlideShape[] = [];
    let notes = "";
    for (const el of elementsOf(treeOf(part))) {
      if (el.name === "p:sp") {
        const spPr = childOf(el, "p:spPr");
        const frame = frameFrom(childOf(spPr, "a:xfrm")) ?? inherited(el);
        if (!frame) warnings.add("a shape without a position was placed at the top left corner");
        const placed = frame ?? { x: 0, y: 0, width: 0, height: 0 };
        const geometry = childOf(spPr, "a:prstGeom")?.attrs.prst;
        // A shape that holds text is a text box here: the model has no text in a shape, and LibreOffice drops the txBox flag.
        const hasText = elementsOf(childOf(el, "p:txBody")).some((p) => p.name === "a:p" && textOf(p).trim() !== "");
        const isText = childOf(childOf(el, "p:nvSpPr"), "p:cNvSpPr")?.attrs.txBox === "1" || phKey(el) !== undefined || hasText;
        if (isText) shapes.push({ type: "textbox", frame: placed, paragraphs: paragraphsOf(el) });
        else if (geometry === "rect" || geometry === "ellipse") {
          const fill = hex(spPr);
          const stroke = hex(childOf(spPr, "a:ln"));
          shapes.push({ type: "shape", kind: geometry === "rect" ? "rectangle" : "ellipse", frame: placed, ...(fill ? { fill } : {}), ...(stroke ? { stroke } : {}) });
        } else if (geometry === "line" && frame) {
          const stroke = hex(childOf(spPr, "a:ln"));
          shapes.push({ type: "shape", kind: "line", frame, ...(stroke ? { stroke } : {}) });
        } else warn(`a ${geometry ?? "custom"} shape`);
      } else if (el.name === "p:cxnSp") {
        const spPr = childOf(el, "p:spPr");
        const frame = frameFrom(childOf(spPr, "a:xfrm"));
        const stroke = hex(childOf(spPr, "a:ln"));
        if (frame) shapes.push({ type: "shape", kind: "line", frame, ...(stroke ? { stroke } : {}) });
        else warn("a connector without a position");
      } else if (el.name === "p:pic") {
        const blip = childOf(childOf(el, "p:blipFill"), "a:blip");
        const target = blip?.attrs["r:embed"] ? slideRels.get(blip.attrs["r:embed"]) : undefined;
        const data = target && !target.external ? pkg.files[target.target] : undefined;
        const sniffed = data ? sniffImage(data) : undefined;
        const frame = frameFrom(childOf(childOf(el, "p:spPr"), "a:xfrm"));
        const alt = childOf(childOf(el, "p:nvPicPr"), "p:cNvPr")?.attrs.descr?.trim();
        if (blip?.attrs["r:link"]) warnings.add("a remote image: not loaded");
        else if (data && sniffed && frame) {
          chargeImages(data.length);
          shapes.push({ type: "picture", frame, data, mime: sniffed.mime, ...(alt ? { alt } : {}) });
        } else warnings.add("an image that is not a PNG, JPEG, or GIF: not imported");
      } else if (el.name === "p:grpSp" || el.name === "p:graphicFrame") warn(el.name === "p:grpSp" ? "a group of shapes" : "a table, chart, or diagram");
    }
    const notesPart = kindTarget(slideRels, "notesSlide");
    if (notesPart && readParts.has(notesPart)) warn("a notes page that two slides share");
    else {
      if (notesPart) readParts.add(notesPart);
      const body = elementsOf(treeOf(notesPart)).find((sp) => sp.name === "p:sp" && phKey(sp)?.startsWith("body|"));
      if (body) notes = paragraphsOf(body).map((runs) => runs.map((r) => r.text).join("")).join("\n");
    }
    slides.push({ shapes, ...(notes ? { notes } : {}) });
  }
  if (!slides.length) throw new FormatError("INVALID_FILE", "the PPTX file has no slide");
  return { model: { width: tenth(width), height: tenth(height), slides }, warnings: [...warnings] };
}
