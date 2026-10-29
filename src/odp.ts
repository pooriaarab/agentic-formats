import { MAX_POINTS, checkDeck, tenth, textBound } from "./deck.js";
import { cm, lengthPt, sniffImage } from "./image.js";
import { FormatError, capped, withLimits, type Limits } from "./limits.js";
import { mergeMarks, normalizeInline, type Deck, type Frame, type Inline, type Mark, type ReadResult, type Slide, type SlideShape, type TextRun, type WriteResult } from "./model.js";
import { ODF_NS_ATTRS, XML_HEADER, readOdfPackage, writeOdfPackage, type OdfImage } from "./odf.js";
import { marksOfProps, runXml, spanStylesXml } from "./odt.js";
import { childOf, elementsOf, safeHref, textOf, xmlEscape as esc, type XmlElement } from "./xml.js";

export const ODP_MIME = "application/vnd.oasis.opendocument.presentation";
const READABLE_MIMES = [ODP_MIME, `${ODP_MIME}-template`];

// ---------- write ----------

interface Out {
  images: OdfImage[];
  spans: Map<string, string>;
  shapeStyles: Map<string, string>;
  warnings: Set<string>;
  charge: (chars: number) => void;
  /** Counts the XML as each piece is made, so a model that holds one shape many times fails before one string grows past the cap. */
  bytes: (length: number) => void;
}

const at = (f: Frame) => `svg:x="${cm(f.x)}" svg:y="${cm(f.y)}" svg:width="${cm(f.width)}" svg:height="${cm(f.height)}"`;

function shapeStyle(shape: { fill?: string; stroke?: string }, out: Out): string {
  const key = `${shape.fill?.toLowerCase() ?? ""}|${shape.stroke?.toLowerCase() ?? ""}`;
  if (!out.shapeStyles.has(key)) out.shapeStyles.set(key, `gs${out.shapeStyles.size + 1}`);
  return out.shapeStyles.get(key)!;
}

function shapeXml(shape: SlideShape, out: Out): string {
  const xml = shapeXmlOf(shape, out);
  out.bytes(xml.length);
  return xml;
}

function shapeXmlOf(shape: SlideShape, out: Out): string {
  const { frame } = shape;
  switch (shape.type) {
    case "textbox": {
      const paragraphs = shape.paragraphs
        .map((runs) => {
          const p = `<text:p text:style-name="Pdeck">${runs.map((r) => runXml(r, out)).join("")}</text:p>`;
          out.bytes(p.length);
          return p;
        })
        .join("");
      return `<draw:frame draw:style-name="gtb" ${at(frame)}><draw:text-box>${paragraphs}</draw:text-box></draw:frame>`;
    }
    case "picture": {
      const sniffed = sniffImage(shape.data);
      if (!sniffed) {
        out.warnings.add("an image is not a PNG, JPEG, or GIF: not written");
        return "";
      }
      const path = `Pictures/Image${out.images.length + 1}.${sniffed.ext}`;
      out.images.push({ path, mime: sniffed.mime, data: shape.data });
      const alt = shape.alt ? `<svg:desc>${esc(shape.alt)}</svg:desc>` : "";
      return `<draw:frame draw:style-name="gtb" ${at(frame)}><draw:image xlink:href="${path}" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>${alt}</draw:frame>`;
    }
    case "shape": {
      const style = `draw:style-name="${shapeStyle(shape, out)}"`;
      if (shape.kind === "line") return `<draw:line ${style} svg:x1="${cm(frame.x)}" svg:y1="${cm(frame.y)}" svg:x2="${cm(frame.x + frame.width)}" svg:y2="${cm(frame.y + frame.height)}"/>`;
      return `<draw:${shape.kind === "ellipse" ? "ellipse" : "rect"} ${style} ${at(frame)}/>`;
    }
  }
}

/** The paragraphs of a slide's notes, one for each line. A newline grows to a paragraph, so the XML counts as it is made. */
function noteLines(notes: string, out: Out): string {
  out.bytes(textBound(notes, 20));
  const xml = notes.split("\n").map((line) => `<text:p>${esc(line)}</text:p>`).join("");
  return xml;
}

/** Write a deck as an OpenDocument presentation (.odp). */
export function writeOdp(deck: Deck, limits?: Partial<Limits>): WriteResult {
  const caps = withLimits(limits);
  checkDeck(deck, caps);
  const out: Out = {
    images: [],
    spans: new Map(),
    shapeStyles: new Map(),
    warnings: new Set(),
    charge: capped(caps.maxTextChars, `the deck has more than ${caps.maxTextChars} characters of text`),
    bytes: capped(caps.maxTotalBytes, `the deck passes ${caps.maxTotalBytes} bytes while it is written`),
  };
  const pages = deck.slides.map((slide, i) => {
    const notes = slide.notes
      ? `<presentation:notes draw:style-name="dp1"><draw:page-thumbnail draw:layer="layout" draw:page-number="${i + 1}" presentation:class="page" ${at({ x: 0, y: 0, width: deck.width, height: deck.height })}/>` +
        `<draw:frame presentation:style-name="prn" draw:layer="layout" presentation:class="notes" presentation:user-transformed="true" ${at({ x: 36, y: deck.height + 36, width: deck.width - 72, height: deck.height })}><draw:text-box>${noteLines(slide.notes, out)}</draw:text-box></draw:frame></presentation:notes>`
      : "";
    return `<draw:page draw:name="page${i + 1}" draw:style-name="dp1" draw:master-page-name="Default">${slide.shapes.map((s) => shapeXml(s, out)).join("")}${notes}</draw:page>`;
  });
  const graphic = [...out.shapeStyles].map(([key, name]) => {
    const [fill, stroke] = key.split("|");
    return `<style:style style:name="${name}" style:family="graphic"><style:graphic-properties ${fill ? `draw:fill="solid" draw:fill-color="${fill}"` : 'draw:fill="none"'} ${stroke ? `draw:stroke="solid" svg:stroke-color="${stroke}" svg:stroke-width="1pt"` : 'draw:stroke="none"'}/></style:style>`;
  });
  const content =
    `${XML_HEADER}<office:document-content ${ODF_NS_ATTRS} office:version="1.3"><office:automatic-styles>` +
    `<style:style style:name="dp1" style:family="drawing-page"/>` +
    `<style:style style:name="gtb" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="none" draw:auto-grow-height="false" draw:auto-grow-width="false" fo:wrap-option="wrap"/></style:style>` +
    `<style:style style:name="prn" style:family="presentation"><style:graphic-properties draw:fill="none" draw:stroke="none"/></style:style>` +
    `<style:style style:name="Pdeck" style:family="paragraph"><style:text-properties fo:font-size="18pt"/></style:style>` +
    `${graphic.join("")}${spanStylesXml(out.spans)}</office:automatic-styles><office:body><office:presentation>${pages.join("")}</office:presentation></office:body></office:document-content>`;
  // LibreOffice takes the slide size from the page layout only when styles.xml also has an office:styles element.
  const styles =
    `${XML_HEADER}<office:document-styles ${ODF_NS_ATTRS} office:version="1.3"><office:styles/><office:automatic-styles>` +
    `<style:page-layout style:name="PM1"><style:page-layout-properties fo:margin-top="0cm" fo:margin-bottom="0cm" fo:margin-left="0cm" fo:margin-right="0cm" fo:page-width="${cm(deck.width)}" fo:page-height="${cm(deck.height)}" style:print-orientation="landscape"/></style:page-layout>` +
    `</office:automatic-styles><office:master-styles><style:master-page style:name="Default" style:page-layout-name="PM1"/></office:master-styles></office:document-styles>`;
  return { bytes: writeOdfPackage({ mime: ODP_MIME, content, styles, images: out.images }, caps), warnings: [...out.warnings] };
}

// ---------- read ----------

const SPACE = "\u0001";
/** A `<text:s>` element stands for this many spaces at most. */
const MAX_SPACES = 1000;

/** Read an OpenDocument presentation (.odp). Content with no place in the model is listed in `warnings`. */
export function readOdp(bytes: Uint8Array, limits?: Partial<Limits>): ReadResult<Deck> {
  const caps = withLimits(limits);
  const { files, content, styles: stylesPart } = readOdfPackage(bytes, caps, "ODP", READABLE_MIMES);
  const presentation = childOf(childOf(content, "office:body"), "office:presentation");
  if (!presentation) throw new FormatError("INVALID_FILE", "the ODP file has no presentation body");
  const warnings = new Set<string>();
  const warn = (what: string) => warnings.add(`${what}: not imported`);
  // A few bytes of `<text:s text:c="1000"/>` stand for a thousand spaces, and an image part is shared by every frame that names it.
  const chargeSpaces = capped(caps.maxTextChars, `the file has more than ${caps.maxTextChars} spaces in <text:s> elements`);
  const chargeImages = capped(caps.maxTotalBytes, `the images in the file add up to more than ${caps.maxTotalBytes} bytes`);

  const styles = new Map<string, { parent?: string; props: Record<string, string>; text: Mark[]; auto: boolean }>();
  const collect = (el: XmlElement, auto: boolean): void => {
    const name = el.attrs["style:name"];
    if (el.name === "style:style" && name) {
      const props = { ...(childOf(el, "style:graphic-properties")?.attrs ?? {}) };
      styles.set(name, { parent: el.attrs["style:parent-style-name"], props, text: marksOfProps(childOf(el, "style:text-properties")?.attrs ?? {}), auto });
    }
    for (const k of elementsOf(el)) collect(k, auto || k.name === "office:automatic-styles");
  };
  collect(content, false);
  if (stylesPart) collect(stylesPart, false);
  /** A style and its parents, nearest first. A loop ends the walk; a chain past `maxStyleDepth` rejects the file. */
  const chain = (name?: string): string[] => {
    const names: string[] = [];
    const seen = new Set<string>();
    for (let n = name; n && !seen.has(n); n = styles.get(n)?.parent) {
      if (names.length >= caps.maxStyleDepth) throw new FormatError("LIMIT_EXCEEDED", `a style has more than ${caps.maxStyleDepth} parents`);
      names.push(n);
      seen.add(n);
    }
    return names;
  };
  const prop = (name: string | undefined, key: string): string | undefined => chain(name).map((n) => styles.get(n)?.props[key]).find((v) => v !== undefined);
  const marksOf = (name?: string): Mark[] => mergeMarks(chain(name).flatMap((n) => styles.get(n)?.text ?? []), []);
  const paragraphMarks = (name?: string): Mark[] => mergeMarks(chain(name).flatMap((n) => (styles.get(n)?.auto ? styles.get(n)!.text : [])), []);
  const color = (name: string | undefined, fill: boolean): string | undefined => {
    const [mode, value] = fill ? [prop(name, "draw:fill"), prop(name, "draw:fill-color")] : [prop(name, "draw:stroke"), prop(name, "svg:stroke-color")];
    return mode === "solid" && value && /^#[0-9a-f]{6}$/i.test(value) ? value.toLowerCase() : undefined;
  };

  const runs = (el: XmlElement, marks: Mark[], href?: string): Inline[] =>
    el.children.flatMap((c): Inline[] => {
      // A viewer underlines a link by itself, so a link run carries no underline mark.
      const run = (text: string, extra: Mark[] = marks): Inline[] => {
        const shown = href ? extra.filter((m) => m !== "underline") : extra;
        return text ? [{ type: "text", text, ...(shown.length ? { marks: shown } : {}), ...(href ? { href } : {}) }] : [];
      };
      if (typeof c === "string") return run(c.replace(/[ \t\r\n]+/g, " "));
      switch (c.name) {
        case "text:s": {
          const asked = Number(c.attrs["text:c"] ?? 1) || 1;
          if (asked > MAX_SPACES) warnings.add(`a run of more than ${MAX_SPACES} spaces was shortened`);
          const count = Math.max(1, Math.min(MAX_SPACES, asked));
          chargeSpaces(count);
          return run(SPACE.repeat(count));
        }
        case "text:tab":
          return run("\t");
        case "text:line-break":
          return run("\n");
        case "text:span":
          return runs(c, mergeMarks(marks, marksOf(c.attrs["text:style-name"])), href);
        case "text:a": {
          const target = safeHref(c.attrs["xlink:href"]);
          if (!target) warnings.add("a link with an unsafe target was read as plain text");
          return runs(c, marks, target ?? href);
        }
        default:
          return runs(c, marks, href);
      }
    });
  const paragraphOf = (p: XmlElement): TextRun[] => {
    const all = runs(p, paragraphMarks(p.attrs["text:style-name"]));
    const first = all.findIndex((n) => n.type === "text");
    const last = all.map((n) => n.type).lastIndexOf("text");
    const trimmed = all.map((n, i) => (n.type !== "text" ? n : { ...n, text: i === first ? (i === last ? n.text.trim() : n.text.trimStart()) : i === last ? n.text.trimEnd() : n.text }));
    return normalizeInline(trimmed.map((n) => (n.type === "text" ? { ...n, text: n.text.replaceAll(SPACE, " ") } : n))) as TextRun[];
  };
  const paragraphsOf = (box: XmlElement): XmlElement[] => elementsOf(box).filter((p) => p.name === "text:p" || p.name === "text:h");

  const length = (el: XmlElement, ...names: string[]): number => {
    for (const name of names) {
      const n = lengthPt(el.attrs[name]);
      if (n !== undefined) return n;
    }
    return 0;
  };
  const frameOf = (el: XmlElement): Frame => {
    const frame = { x: length(el, "svg:x"), y: length(el, "svg:y"), width: length(el, "svg:width"), height: length(el, "svg:height", "fo:min-height") };
    // A position or size past the bounds the writers enforce is not read.
    const usable = [frame.x, frame.y].every((n) => Math.abs(n) <= MAX_POINTS) && [frame.width, frame.height].every((n) => n >= 0 && n <= MAX_POINTS);
    if (usable) return frame;
    warnings.add("a shape with a position or size outside the slide bounds was placed at the top left corner");
    return { x: 0, y: 0, width: 0, height: 0 };
  };

  function shapesOf(page: XmlElement): SlideShape[] {
    const out: SlideShape[] = [];
    for (const el of elementsOf(page)) {
      const style = el.attrs["draw:style-name"];
      switch (el.name) {
        case "draw:frame": {
          const box = childOf(el, "draw:text-box");
          const picture = childOf(el, "draw:image");
          if (box) out.push({ type: "textbox", frame: frameOf(el), paragraphs: paragraphsOf(box).map(paragraphOf) });
          else if (picture) {
            const href = picture.attrs["xlink:href"] ?? "";
            const data = files[href];
            const sniffed = data ? sniffImage(data) : undefined;
            const alt = textOf(childOf(el, "svg:desc")).trim() || textOf(childOf(el, "svg:title")).trim();
            if (data && sniffed) {
              chargeImages(data.length);
              out.push({ type: "picture", frame: frameOf(el), data, mime: sniffed.mime, ...(alt ? { alt } : {}) });
            } else warnings.add(/^[a-z][a-z0-9+.-]*:/i.test(href) ? "a remote image: not loaded" : "an image that is not a PNG, JPEG, or GIF: not imported");
          } else warn(el.name);
          break;
        }
        case "draw:rect":
        case "draw:ellipse":
        case "draw:circle":
        case "draw:custom-shape": {
          const preset = el.name === "draw:custom-shape" ? childOf(el, "draw:enhanced-geometry")?.attrs["draw:type"] : el.name === "draw:rect" ? "rectangle" : "ellipse";
          // LibreOffice saves a text box from a PPTX as a custom shape that holds the text.
          const text = paragraphsOf(el).map(paragraphOf);
          if (text.some((runs) => runs.length)) {
            out.push({ type: "textbox", frame: frameOf(el), paragraphs: text });
            break;
          }
          const type = preset === "ooxml-rect" ? "rectangle" : preset === "ooxml-ellipse" ? "ellipse" : preset;
          if (type !== "rectangle" && type !== "ellipse") {
            warn(`a ${preset ?? "custom"} shape`);
            break;
          }
          const fill = color(style, true);
          const stroke = color(style, false);
          out.push({ type: "shape", kind: type, frame: frameOf(el), ...(fill ? { fill } : {}), ...(stroke ? { stroke } : {}) });
          break;
        }
        case "draw:line":
        case "draw:connector": {
          const [x1, y1, x2, y2] = ["svg:x1", "svg:y1", "svg:x2", "svg:y2"].map((a) => length(el, a));
          const stroke = color(style, false);
          out.push({ type: "shape", kind: "line", frame: { x: tenth(Math.min(x1, x2)), y: tenth(Math.min(y1, y2)), width: tenth(Math.abs(x2 - x1)), height: tenth(Math.abs(y2 - y1)) }, ...(stroke ? { stroke } : {}) });
          break;
        }
        case "presentation:notes":
        case "office:forms":
        case "presentation:animations":
          break;
        default:
          warn(el.name);
      }
    }
    return out;
  }

  const slides: Slide[] = elementsOf(presentation)
    .filter((page) => page.name === "draw:page")
    .map((page) => {
      const box = elementsOf(childOf(page, "presentation:notes")).map((f) => childOf(f, "draw:text-box")).find(Boolean);
      const notes = box ? paragraphsOf(box).map((p) => paragraphOf(p).map((r) => r.text).join("")).join("\n") : "";
      return { shapes: shapesOf(page), ...(notes ? { notes } : {}) };
    });
  if (!slides.length) throw new FormatError("INVALID_FILE", "the ODP file has no slide");
  // The size is the layout of the master page the first slide uses. A file also has layouts for its notes pages.
  const masterName = elementsOf(presentation).find((page) => page.name === "draw:page")?.attrs["draw:master-page-name"];
  const masterPage = [content, stylesPart].flatMap((part) => (part ? [...allOf(part, "style:master-page")] : [])).find((m) => m.attrs["style:name"] === masterName);
  const layout = [content, stylesPart]
    .flatMap((part) => (part ? [...allOf(part, "style:page-layout")] : []))
    .find((l) => l.attrs["style:name"] === masterPage?.attrs["style:page-layout-name"]);
  const layoutProps = childOf(layout, "style:page-layout-properties");
  const width = lengthPt(layoutProps?.attrs["fo:page-width"]);
  const height = lengthPt(layoutProps?.attrs["fo:page-height"]);
  const sized = width !== undefined && height !== undefined && width > 0 && height > 0 && width <= MAX_POINTS && height <= MAX_POINTS;
  if (!sized) warnings.add("the slide size was missing: 720 by 405 points assumed");
  return { model: { width: sized ? width : 720, height: sized ? height : 405, slides }, warnings: [...warnings] };
}

function* allOf(el: XmlElement, name: string): Generator<XmlElement> {
  if (el.name === name) yield el;
  for (const k of elementsOf(el)) yield* allOf(k, name);
}
