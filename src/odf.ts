import { strFromU8 } from "fflate";
import { FormatError, type Limits } from "./limits.js";
import { childOf, parseXml, textOf, xmlEscape, type Namespaces, type XmlElement } from "./xml.js";
import { readZip, writeZip, type ZipEntry } from "./zip.js";

/** The namespaces of ODF, each mapped to the prefix the spec uses. */
export const ODF_NAMESPACES: Namespaces = {
  "urn:oasis:names:tc:opendocument:xmlns:office:1.0": "office",
  "urn:oasis:names:tc:opendocument:xmlns:text:1.0": "text",
  "urn:oasis:names:tc:opendocument:xmlns:table:1.0": "table",
  "urn:oasis:names:tc:opendocument:xmlns:style:1.0": "style",
  "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0": "fo",
  "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0": "draw",
  "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0": "presentation",
  "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0": "svg",
  "urn:oasis:names:tc:opendocument:xmlns:meta:1.0": "meta",
  "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0": "number",
  "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0": "manifest",
  "http://www.w3.org/1999/xlink": "xlink",
  "http://purl.org/dc/elements/1.1/": "dc",
};

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8"?>\n';

/** Namespace declarations for the root of a content, styles, or meta part. */
export const ODF_NS_ATTRS = [
  ["office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0"],
  ["text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0"],
  ["table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"],
  ["style", "urn:oasis:names:tc:opendocument:xmlns:style:1.0"],
  ["fo", "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"],
  ["xlink", "http://www.w3.org/1999/xlink"],
  ["draw", "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"],
  ["presentation", "urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"],
  ["svg", "urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"],
  ["meta", "urn:oasis:names:tc:opendocument:xmlns:meta:1.0"],
  ["number", "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"],
  ["dc", "http://purl.org/dc/elements/1.1/"],
]
  .map(([prefix, uri]) => `xmlns:${prefix}="${uri}"`)
  .join(" ");

/** Text for an ODF paragraph: tabs, line breaks, and runs of spaces become their elements. */
export const spaced = (s: string): string =>
  xmlEscape(s)
    .replace(/\t/g, "<text:tab/>")
    .replace(/\n/g, "<text:line-break/>")
    .replace(/ +/g, (run, at: number, all: string) =>
      at === 0 || at + run.length === all.length ? `<text:s text:c="${run.length}"/>` : run.length > 1 ? ` <text:s text:c="${run.length - 1}"/>` : run,
    );

/**
 * The most characters `spaced(text)` can make, found without making them. A newline becomes 18, a tab 11, a
 * quote or an angle bracket up to 6, and a space next to another space or at an end up to 20.
 */
export function spacedBound(text: string): number {
  let n = text.length;
  let previous = " ";
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (c === "\n" || c === "\t" || c === "&" || c === "<" || c === ">" || c === '"' || c === "\r" || (c === " " && (previous === " " || i === text.length - 1))) n += 20;
    previous = c;
  }
  return n;
}

export interface OdfImage {
  /** Path in the package, such as `Pictures/Image1.png`. */
  path: string;
  mime: string;
  data: Uint8Array;
}

export interface OdfParts {
  mime: string;
  /** The whole `content.xml`. */
  content: string;
  /** The whole `styles.xml`. */
  styles: string;
  title?: string;
  images: OdfImage[];
}

/** Pack an ODF file. `mimetype` is first and stored, so a file sniffer can read it. */
export function writeOdfPackage(parts: OdfParts, limits: Limits): Uint8Array {
  const entry = (path: string, mime: string) => `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="${mime}"/>`;
  const manifest =
    `${XML_HEADER}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="${parts.mime}"/>` +
    ["content.xml", "styles.xml", "meta.xml"].map((n) => entry(n, "text/xml")).join("") +
    parts.images.map((i) => entry(i.path, i.mime)).join("") +
    `</manifest:manifest>`;
  const meta =
    `${XML_HEADER}<office:document-meta ${ODF_NS_ATTRS} office:version="1.3"><office:meta><meta:generator>agentic-formats</meta:generator>` +
    `${parts.title ? `<dc:title>${xmlEscape(parts.title)}</dc:title>` : ""}</office:meta></office:document-meta>`;
  const entries: ZipEntry[] = [
    { name: "mimetype", data: parts.mime, store: true },
    { name: "META-INF/manifest.xml", data: manifest },
    { name: "content.xml", data: parts.content },
    { name: "styles.xml", data: parts.styles },
    { name: "meta.xml", data: meta },
    ...parts.images.map((i) => ({ name: i.path, data: i.data })),
  ];
  return writeZip(entries, limits);
}

export interface OdfFile {
  files: Record<string, Uint8Array>;
  content: XmlElement;
  styles?: XmlElement;
  title?: string;
}

/** Unpack an ODF file and parse its parts. `mimes` lists the media types the caller accepts. */
export function readOdfPackage(bytes: Uint8Array, limits: Limits, kind: string, mimes: string[]): OdfFile {
  const files = readZip(bytes, limits, kind, (n) => /^(mimetype|content\.xml|styles\.xml|meta\.xml|Pictures\/.+)$/.test(n));
  const mime = files.mimetype ? strFromU8(files.mimetype).trim() : "";
  if (!mimes.includes(mime)) throw new FormatError("INVALID_FILE", `the file is not an ${kind} file`);
  if (!files["content.xml"]) throw new FormatError("INVALID_FILE", `the ${kind} file has no content.xml`);
  const parse = (part: string) => parseXml(files[part], limits, ODF_NAMESPACES);
  const meta = files["meta.xml"] ? parse("meta.xml") : undefined;
  const title = textOf(childOf(childOf(meta, "office:meta"), "dc:title")).trim();
  return { files, content: parse("content.xml"), styles: files["styles.xml"] ? parse("styles.xml") : undefined, ...(title ? { title } : {}) };
}
