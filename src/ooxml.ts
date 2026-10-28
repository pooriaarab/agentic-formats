import { FormatError, type Limits } from "./limits.js";
import { elementsOf, parseXml, xmlEscape as esc, type Namespaces, type XmlElement } from "./xml.js";
import { isSafeEntryName, readZip } from "./zip.js";

const NS = (prefix: string, ...uris: string[]): [string, string][] => uris.map((uri) => [uri, prefix]);
const ECMA = "http://schemas.openxmlformats.org";
const STRICT = "http://purl.oclc.org/ooxml";

/** OOXML namespaces, each mapped to a fixed prefix. The strict variants map to the same prefixes. */
export const OOXML_NAMESPACES: Namespaces = Object.fromEntries([
  ...NS("", `${ECMA}/spreadsheetml/2006/main`, `${STRICT}/spreadsheetml/main`, `${ECMA}/package/2006/relationships`, `${ECMA}/package/2006/content-types`),
  ...NS("w", `${ECMA}/wordprocessingml/2006/main`, `${STRICT}/wordprocessingml/main`),
  ...NS("a", `${ECMA}/drawingml/2006/main`, `${STRICT}/drawingml/main`),
  ...NS("p", `${ECMA}/presentationml/2006/main`, `${STRICT}/presentationml/main`),
  ...NS("r", `${ECMA}/officeDocument/2006/relationships`, `${STRICT}/officeDocument/relationships`),
  ...NS("wp", `${ECMA}/drawingml/2006/wordprocessingDrawing`, `${STRICT}/drawingml/wordprocessingDrawing`),
  ...NS("pic", `${ECMA}/drawingml/2006/picture`, `${STRICT}/drawingml/picture`),
  ...NS("cp", `${ECMA}/package/2006/metadata/core-properties`),
  ...NS("dc", "http://purl.org/dc/elements/1.1/"),
]);

export const XML_HEADER = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
export const REL_NS = `${ECMA}/officeDocument/2006/relationships`;
export const RELS_NS = `${ECMA}/package/2006/relationships`;
const CONTENT_TYPES_NS = `${ECMA}/package/2006/content-types`;

export interface Relationship {
  id: string;
  /** The last part of the relationship type URI: `worksheet`, `image`, `hyperlink`. */
  kind: string;
  /** A package path for an internal target; the URI for an external one. */
  target: string;
  external: boolean;
}

/** `[Content_Types].xml` from default types by extension and override types by part. */
export function contentTypesXml(defaults: Record<string, string>, overrides: Record<string, string>): string {
  return (
    `${XML_HEADER}<Types xmlns="${CONTENT_TYPES_NS}">` +
    Object.entries(defaults).map(([ext, type]) => `<Default Extension="${ext}" ContentType="${type}"/>`).join("") +
    Object.entries(overrides).map(([part, type]) => `<Override PartName="/${part}" ContentType="${type}"/>`).join("") +
    `</Types>`
  );
}

/** A `.rels` part. `type` is the last part of the type URI. */
export function relsXml(rels: { id: string; type: string; target: string; external?: boolean }[]): string {
  return (
    `${XML_HEADER}<Relationships xmlns="${RELS_NS}">` +
    rels.map((r) => `<Relationship Id="${r.id}" Type="${REL_NS}/${r.type}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ""}/>`).join("") +
    `</Relationships>`
  );
}

/** Join a relationship target to the part that holds it. Returns undefined for a path that leaves the package. */
export function resolveTarget(source: string, target: string): string | undefined {
  const parts = target.startsWith("/") ? [] : source.split("/").slice(0, -1);
  for (const part of target.split("/")) {
    if (part === "..") {
      if (parts.pop() === undefined) return undefined;
    } else if (part !== "." && part !== "") parts.push(part);
  }
  const path = parts.join("/");
  return isSafeEntryName(path) ? path : undefined;
}

export interface OoxmlFile {
  files: Record<string, Uint8Array>;
  parse: (part: string) => XmlElement;
  /** The relationships of a part, by id. External targets are listed, never followed. */
  rels: (part: string) => Map<string, Relationship>;
  /** The main part, found through the root relationships. */
  mainPart: string;
}

/** Unpack an OOXML file, then find its main part. `mainKind` is `officeDocument` for all three formats. */
export function readOoxmlPackage(bytes: Uint8Array, limits: Limits, kind: string, mainContentType: string): OoxmlFile {
  const files = readZip(bytes, limits, kind, (n) => /\.(xml|rels|png|jpe?g|gif)$/i.test(n));
  const parse = (part: string): XmlElement => {
    if (!files[part]) throw new FormatError("INVALID_FILE", `the ${kind} file has no ${part}`);
    return parseXml(files[part], limits, OOXML_NAMESPACES);
  };
  const rels = (part: string): Map<string, Relationship> => {
    const dir = part.split("/").slice(0, -1).join("/");
    const path = `${dir ? `${dir}/` : ""}_rels/${part.split("/").pop()}.rels`;
    const out = new Map<string, Relationship>();
    if (!files[path]) return out;
    for (const r of elementsOf(parse(path))) {
      const external = r.attrs.TargetMode === "External";
      const target = external ? r.attrs.Target : resolveTarget(part, r.attrs.Target ?? "");
      if (r.attrs.Id && r.attrs.Type && target !== undefined) out.set(r.attrs.Id, { id: r.attrs.Id, kind: r.attrs.Type.split("/").pop()!, target, external });
    }
    return out;
  };
  const root = [...rels("")].find(([, r]) => r.kind === "officeDocument" && !r.external)?.[1];
  const mainPart = root?.target;
  if (!mainPart) throw new FormatError("INVALID_FILE", `the file is not an ${kind} file`);
  const types = files["[Content_Types].xml"] ? elementsOf(parse("[Content_Types].xml")) : [];
  const type = types.find((t) => t.name === "Override" && t.attrs.PartName === `/${mainPart}`)?.attrs.ContentType ?? "";
  if (!type.startsWith(mainContentType)) throw new FormatError("INVALID_FILE", `the file is not an ${kind} file`);
  return { files, parse, rels, mainPart };
}


