import { cm, frameSize, lengthPt, sniffImage } from "./image.js";
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
  type ListItem,
  type Mark,
  type ReadResult,
  type Table,
  type TextDoc,
  type TextRun,
  type WriteResult,
} from "./model.js";
import { ODF_NS_ATTRS, XML_HEADER, readOdfPackage, spaced, writeOdfPackage, type OdfImage } from "./odf.js";
import { childOf, elementsOf, safeHref, textOf, xmlEscape as esc, type XmlElement } from "./xml.js";

export const ODT_MIME = "application/vnd.oasis.opendocument.text";
/** Writer saves a web page or a template with a sibling type. The body format is the same. */
const READABLE_MIMES = [ODT_MIME, `${ODT_MIME}-web`, `${ODT_MIME}-template`];

// ---------- write ----------

const MARK_PROPS: Record<Mark, string> = {
  bold: 'fo:font-weight="bold" style:font-weight-asian="bold" style:font-weight-complex="bold"',
  italic: 'fo:font-style="italic" style:font-style-asian="italic" style:font-style-complex="italic"',
  underline: 'style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"',
  strike: 'style:text-line-through-style="solid"',
  code: 'fo:font-family="\'Liberation Mono\', \'Courier New\', monospace" style:font-family-generic="modern"',
  superscript: 'style:text-position="super 58%"',
  subscript: 'style:text-position="sub 58%"',
};

const LIST_LEVELS = (ordered: boolean): string =>
  Array.from({ length: 10 }, (_, i) => {
    const place = `<style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="listtab" fo:margin-left="${((i + 1) * 0.63).toFixed(2)}cm" fo:text-indent="-0.63cm"/></style:list-level-properties>`;
    return ordered
      ? `<text:list-level-style-number text:level="${i + 1}" style:num-suffix="." style:num-format="1">${place}</text:list-level-style-number>`
      : `<text:list-level-style-bullet text:level="${i + 1}" text:bullet-char="•">${place}</text:list-level-style-bullet>`;
  }).join("");

const HEADING_SIZES = [24, 18, 14, 12, 11, 10];
const STYLES =
  `${XML_HEADER}<office:document-styles ${ODF_NS_ATTRS} office:version="1.3"><office:styles>` +
  `<style:style style:name="Standard" style:family="paragraph" style:class="text"/>` +
  HEADING_SIZES.map(
    (size, i) =>
      `<style:style style:name="Heading_20_${i + 1}" style:display-name="Heading ${i + 1}" style:family="paragraph" style:parent-style-name="Standard" style:next-style-name="Standard" style:default-outline-level="${i + 1}" style:class="text"><style:text-properties fo:font-size="${size}pt" fo:font-weight="bold"/></style:style>`,
  ).join("") +
  `<style:style style:name="Footnote" style:family="paragraph" style:parent-style-name="Standard" style:class="extra"><style:text-properties fo:font-size="10pt"/></style:style>` +
  `</office:styles></office:document-styles>`;

interface Out {
  images: OdfImage[];
  spans: Map<string, string>;
  tables: number;
  notes: number;
  warnings: Set<string>;
}

function imageXml(node: InlineImage, out: Out): string {
  const sniffed = sniffImage(node.data);
  if (!sniffed) {
    out.warnings.add("an image is not a PNG, JPEG, or GIF: not written");
    return "";
  }
  const name = `Image${out.images.length + 1}`;
  const path = `Pictures/${name}.${sniffed.ext}`;
  out.images.push({ path, mime: sniffed.mime, data: node.data });
  const [w, h] = frameSize(node.width, node.height, sniffed);
  const alt = node.alt ? `<svg:title>${esc(node.alt)}</svg:title>` : "";
  return `<draw:frame draw:name="${name}" text:anchor-type="as-char" svg:width="${cm(w)}" svg:height="${cm(h)}"><draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>${alt}</draw:frame>`;
}

function runXml(run: TextRun, out: Out): string {
  let xml = spaced(run.text);
  const marks = MARKS.filter((m) => run.marks?.includes(m));
  if (marks.length) {
    const key = marks.join(",");
    if (!out.spans.has(key)) out.spans.set(key, `T${out.spans.size + 1}`);
    xml = `<text:span text:style-name="${out.spans.get(key)}">${xml}</text:span>`;
  }
  if (run.href !== undefined) {
    const href = safeHref(run.href);
    if (href) xml = `<text:a xlink:type="simple" xlink:href="${esc(href)}">${xml}</text:a>`;
    else out.warnings.add("a link with an unsafe target was written as plain text");
  }
  return xml;
}

function inlineXml(nodes: Inline[], out: Out): string {
  return nodes
    .map((node) => {
      switch (node.type) {
        case "text":
          return runXml(node, out);
        case "break":
          return "<text:line-break/>";
        case "image":
          return imageXml(node, out);
        case "footnote": {
          const n = ++out.notes;
          return `<text:note text:id="ftn${n}" text:note-class="footnote"><text:note-citation>${n}</text:note-citation><text:note-body><text:p text:style-name="Footnote">${inlineXml(node.content, out)}</text:p></text:note-body></text:note>`;
        }
      }
    })
    .join("");
}

function listXml(list: List, out: Out): string {
  const items = list.items.map((item) => `<text:list-item><text:p text:style-name="Standard">${inlineXml(item.content, out)}</text:p>${item.sublist ? listXml(item.sublist, out) : ""}</text:list-item>`);
  return `<text:list text:style-name="${list.ordered ? "Numbering" : "Bullets"}">${items.join("")}</text:list>`;
}

function tableXml(table: Table, out: Out): string {
  const columns = Math.max(0, ...table.rows.map((r) => r.length));
  const row = (cells: Block[][]) =>
    `<table:table-row>${Array.from({ length: columns }, (_, i) => {
      const blocks = cells[i] ?? [];
      return `<table:table-cell office:value-type="string">${blocks.length ? blocks.map((b) => blockXml(b, out)).join("") : `<text:p text:style-name="Standard"/>`}</table:table-cell>`;
    }).join("")}</table:table-row>`;
  const [first, ...rest] = table.rows;
  const rows = table.header && first ? `<table:table-header-rows>${row(first)}</table:table-header-rows>${rest.map(row).join("")}` : table.rows.map(row).join("");
  return `<table:table table:name="Table${++out.tables}">${columns ? `<table:table-column table:number-columns-repeated="${columns}"/>` : ""}${rows}</table:table>`;
}

function blockXml(block: Block, out: Out): string {
  switch (block.type) {
    case "paragraph":
      return `<text:p text:style-name="Standard">${inlineXml(block.content, out)}</text:p>`;
    case "heading": {
      const level = Math.min(6, Math.max(1, Math.trunc(block.level) || 1));
      return `<text:h text:style-name="Heading_20_${level}" text:outline-level="${level}">${inlineXml(block.content, out)}</text:h>`;
    }
    case "list":
      return listXml(block, out);
    case "table":
      return tableXml(block, out);
  }
}

/** Write a document as an OpenDocument text file (.odt). */
export function writeOdt(doc: TextDoc, limits?: Partial<Limits>): WriteResult {
  const caps = withLimits(limits);
  const out: Out = { images: [], spans: new Map(), tables: 0, notes: 0, warnings: new Set() };
  const body = doc.blocks.map((b) => blockXml(b, out)).join("\n");
  const spanStyles = [...out.spans].map(
    ([key, name]) => `<style:style style:name="${name}" style:family="text"><style:text-properties ${key.split(",").map((m) => MARK_PROPS[m as Mark]).join(" ")}/></style:style>`,
  );
  const content =
    `${XML_HEADER}<office:document-content ${ODF_NS_ATTRS} office:version="1.3"><office:automatic-styles>${spanStyles.join("")}` +
    `<text:list-style style:name="Bullets">${LIST_LEVELS(false)}</text:list-style><text:list-style style:name="Numbering">${LIST_LEVELS(true)}</text:list-style>` +
    `</office:automatic-styles><office:body><office:text>${body}</office:text></office:body></office:document-content>`;
  const bytes = writeOdfPackage({ mime: ODT_MIME, content, styles: STYLES, title: doc.title, images: out.images }, caps);
  return { bytes, warnings: [...out.warnings] };
}

// ---------- read ----------

const SPACE = "\u0001"; // a space that came from <text:s>, kept through trimming
const NAMED_MARKS: Record<string, Mark[]> = { Source_20_Text: ["code"], Strong_20_Emphasis: ["bold"], Emphasis: ["italic"] };
const IGNORED = /^(text:(sequence-decls|variable-decls|user-field-decls|tracked-changes|soft-page-break|bookmark|bookmark-start|bookmark-end|index-title|toc-mark.*)|office:forms|text:note-citation)$/;
/** A repeat count above this is a hostile or broken file, not a table. */
const MAX_REPEAT = 1000;

interface StyleInfo {
  parent?: string;
  marks: Mark[];
  /** Made for one paragraph or span, so it can mark all the text it covers. */
  auto: boolean;
}

function marksOfProps(props: Record<string, string>): Mark[] {
  const marks: Mark[] = [];
  if (/^(bold|[6-9]00)$/.test(props["fo:font-weight"] ?? "")) marks.push("bold");
  if (props["fo:font-style"] === "italic") marks.push("italic");
  const underline = props["style:text-underline-style"];
  if (underline && underline !== "none") marks.push("underline");
  const strike = props["style:text-line-through-style"];
  if (strike && strike !== "none") marks.push("strike");
  if (/mono|courier|consolas/i.test(props["fo:font-family"] ?? props["style:font-name"] ?? "")) marks.push("code");
  const position = (props["style:text-position"] ?? "").split(/\s+/)[0];
  if (position === "super" || /^[1-9][0-9.]*%$/.test(position)) marks.push("superscript");
  else if (position === "sub" || /^-[0-9.]+%$/.test(position)) marks.push("subscript");
  return marks;
}

/** Read an OpenDocument text file (.odt). Content with no place in the model is listed in `warnings`. */
export function readOdt(bytes: Uint8Array, limits?: Partial<Limits>): ReadResult<TextDoc> {
  const caps = withLimits(limits);
  const { files, content, styles: stylesPart, title } = readOdfPackage(bytes, caps, "ODT", READABLE_MIMES);
  const text = childOf(childOf(content, "office:body"), "office:text");
  if (!text) throw new FormatError("INVALID_FILE", "the ODT file has no text body");

  const styles = new Map<string, StyleInfo>();
  const listStyles = new Map<string, XmlElement>();
  const collect = (el: XmlElement, auto: boolean): void => {
    const name = el.attrs["style:name"];
    if (el.name === "style:style" && name) {
      const marks = [...(NAMED_MARKS[name] ?? []), ...marksOfProps(childOf(el, "style:text-properties")?.attrs ?? {})];
      styles.set(name, { parent: el.attrs["style:parent-style-name"], marks, auto });
    } else if (el.name === "text:list-style" && name) listStyles.set(name, el);
    for (const k of elementsOf(el)) collect(k, auto || k.name === "office:automatic-styles");
  };
  collect(content, false);
  if (stylesPart) collect(stylesPart, false);

  const chain = (name?: string): string[] => {
    const names: string[] = [];
    for (let n = name; n && !names.includes(n); n = styles.get(n)?.parent) names.push(n);
    return names;
  };
  const marksOf = (name?: string): Mark[] => chain(name).flatMap((n) => styles.get(n)?.marks ?? NAMED_MARKS[n] ?? []);
  const paragraphMarks = (name?: string): Mark[] => chain(name).flatMap((n) => (styles.get(n)?.auto ? styles.get(n)!.marks : []));
  const isOrdered = (styleName: string | undefined, depth: number, inherited: boolean): boolean => {
    const style = styleName ? listStyles.get(styleName) : undefined;
    if (!style) return inherited;
    const levels = elementsOf(style);
    const level = levels.find((l) => l.attrs["text:level"] === String(depth + 1)) ?? levels[0];
    return level?.name === "text:list-level-style-number";
  };

  const warnings = new Set<string>();
  const warn = (what: string) => warnings.add(`${what}: not imported`);
  const run = (value: string, marks: Mark[], href?: string): TextRun => ({ type: "text", text: value, ...(marks.length ? { marks } : {}), ...(href ? { href } : {}) });

  function frame(el: XmlElement): Inline[] {
    const picture = childOf(el, "draw:image");
    if (!picture) return warn(el.name), [];
    const href = picture.attrs["xlink:href"] ?? "";
    const data = files[href];
    const sniffed = data ? sniffImage(data) : undefined;
    if (!data || !sniffed) {
      warnings.add(/^[a-z][a-z0-9+.-]*:/i.test(href) ? "a remote image: not loaded" : "an image that is not a PNG, JPEG, or GIF: not imported");
      return [];
    }
    const alt = textOf(childOf(el, "svg:title")).trim() || textOf(childOf(el, "svg:desc")).trim();
    const width = lengthPt(el.attrs["svg:width"]);
    const height = lengthPt(el.attrs["svg:height"]);
    return [{ type: "image", data, mime: sniffed.mime, ...(alt ? { alt } : {}), ...(width && height ? { width, height } : {}) }];
  }

  const inline = (el: XmlElement, marks: Mark[], href?: string): Inline[] =>
    el.children.flatMap((c): Inline[] => {
      if (typeof c === "string") {
        const value = c.replace(/[ \t\r\n]+/g, " ");
        return value ? [run(value, marks, href)] : [];
      }
      switch (c.name) {
        case "text:s":
          return [run(SPACE.repeat(Math.max(1, Math.min(1000, Number(c.attrs["text:c"] ?? 1)) || 1)), marks, href)];
        case "text:tab":
          return [run("\t", marks, href)];
        case "text:line-break":
          return [{ type: "break" }];
        case "text:span":
          return inline(c, [...marks, ...marksOf(c.attrs["text:style-name"])], href);
        case "text:a": {
          const target = safeHref(c.attrs["xlink:href"]);
          if (!target) warnings.add("a link with an unsafe target was read as plain text");
          return inline(c, marks, target ?? href);
        }
        case "draw:frame":
          return frame(c);
        case "text:note": {
          if (c.attrs["text:note-class"] === "endnote") warnings.add("an endnote was read as a footnote");
          const paragraphs = elementsOf(childOf(c, "text:note-body")).filter((p) => p.name === "text:p" || p.name === "text:h");
          const body = paragraphs.flatMap((p, i): Inline[] => [...(i ? [{ type: "break" } as const] : []), ...finish(inline(p, paragraphMarks(p.attrs["text:style-name"])))]);
          const note: Footnote = { type: "footnote", content: body.filter((n): n is TextRun | { type: "break" } => n.type === "text" || n.type === "break") };
          if (note.content.length !== body.length) warnings.add("an image in a footnote: not imported");
          return [note];
        }
        case "office:annotation":
          return warn(c.name), [];
        default:
          return IGNORED.test(c.name) ? [] : inline(c, marks, href);
      }
    });

  /** Trim the ends of a paragraph, turn the kept spaces back into spaces, and join equal runs. */
  const finish = (nodes: Inline[]): Inline[] => {
    const isText = (n: Inline) => n.type === "text";
    const first = nodes.findIndex(isText);
    const last = nodes.map((n) => n.type).lastIndexOf("text");
    const trimmed = nodes.map((n, i) => {
      if (n.type !== "text") return n;
      const t = n.text;
      return { ...n, text: i === first ? (i === last ? t.trim() : t.trimStart()) : i === last ? t.trimEnd() : t };
    });
    return normalizeInline(trimmed.map((n) => (n.type === "text" ? { ...n, text: n.text.replaceAll(SPACE, " ") } : n)));
  };
  const inlineOf = (el: XmlElement): Inline[] => finish(inline(el, paragraphMarks(el.attrs["text:style-name"])));

  function listOf(el: XmlElement, depth: number, inherited: boolean): List {
    const ordered = isOrdered(el.attrs["text:style-name"], depth, inherited);
    const items = elementsOf(el)
      .filter((k) => k.name === "text:list-item" || k.name === "text:list-header")
      .map((item): ListItem => {
        const kids = elementsOf(item);
        const paragraphs = kids.filter((k) => k.name === "text:p" || k.name === "text:h");
        const lists = kids.filter((k) => k.name === "text:list").map((l) => listOf(l, depth + 1, ordered));
        if (paragraphs.length > 1) warnings.add("a list item with several paragraphs: paragraphs were joined with line breaks");
        if (kids.some((k) => k.name === "table:table")) warn("a table in a list item");
        const content = paragraphs.flatMap((p, i): Inline[] => [...(i ? [{ type: "break" } as const] : []), ...inlineOf(p)]);
        const sublist = lists.length ? { ...lists[0], items: lists.flatMap((l) => l.items) } : undefined;
        return { content, ...(sublist ? { sublist } : {}) };
      });
    return { type: "list", ordered, items };
  }

  const repeat = (el: XmlElement, attr: string): number => {
    const n = Number.parseInt(el.attrs[attr] ?? "1", 10);
    if (!Number.isFinite(n) || n < 1) return 1;
    if (n > MAX_REPEAT) throw new FormatError("LIMIT_EXCEEDED", `a table repeats a row or cell more than ${MAX_REPEAT} times`);
    return n;
  };

  function tableOf(el: XmlElement): Table | undefined {
    const rows: Block[][][] = [];
    let header = false;
    const walk = (parent: XmlElement, inHeader: boolean): void => {
      for (const k of elementsOf(parent)) {
        if (k.name === "table:table-row") {
          const cells = elementsOf(k).flatMap((c): Block[][] => {
            if (c.name !== "table:table-cell" && c.name !== "table:covered-table-cell") return [];
            if (c.name === "table:covered-table-cell" || c.attrs["table:number-columns-spanned"] || c.attrs["table:number-rows-spanned"]) warnings.add("merged table cells were read as single cells");
            const blocks = blocksOf(c, 0, false);
            const empty = blocks.length === 1 && blocks[0].type === "paragraph" && blocks[0].content.length === 0;
            return Array.from({ length: repeat(c, "table:number-columns-repeated") }, () => (empty ? [] : blocks));
          });
          if (inHeader && rows.length === 0) header = true;
          else if (inHeader) warnings.add("only the first header row of a table was kept as a header");
          for (let i = repeat(k, "table:number-rows-repeated"); i > 0; i--) rows.push(cells);
        } else if (/^table:(table-header-rows|table-rows|table-row-group)$/.test(k.name)) walk(k, inHeader || k.name === "table:table-header-rows");
      }
    };
    walk(el, false);
    return rows.length ? { type: "table", ...(header ? { header: true } : {}), rows } : undefined;
  }

  function blocksOf(parent: XmlElement, depth: number, inherited: boolean): Block[] {
    const out: Block[] = [];
    for (const el of elementsOf(parent)) {
      const names = chain(el.attrs["text:style-name"]);
      switch (el.name) {
        case "text:h":
          out.push({ type: "heading", level: Math.min(6, Math.max(1, Number.parseInt(el.attrs["text:outline-level"] ?? "1", 10) || 1)) as Heading["level"], content: inlineOf(el) });
          break;
        case "text:p":
          out.push(names.includes("Title") ? { type: "heading", level: 1, content: inlineOf(el) } : { type: "paragraph", content: inlineOf(el) });
          break;
        case "text:list":
          out.push(listOf(el, depth, inherited));
          break;
        case "table:table": {
          const table = tableOf(el);
          if (table) out.push(table);
          break;
        }
        case "draw:frame":
          out.push({ type: "paragraph", content: frame(el) });
          break;
        case "text:section":
          out.push(...blocksOf(el, depth, inherited));
          break;
        default:
          if (!IGNORED.test(el.name)) warn(el.name);
      }
    }
    return out;
  }

  const doc: TextDoc = { ...(title ? { title } : {}), blocks: blocksOf(text, 0, false) };
  return { model: normalizeTextDoc(doc), warnings: [...warnings] };
}
