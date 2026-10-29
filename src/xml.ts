import { strFromU8 } from "fflate";
import { FormatError, type Limits } from "./limits.js";

export interface XmlElement {
  /** Canonical `prefix:local` for a known namespace, the name as written for any other. */
  name: string;
  attrs: Record<string, string>;
  children: (XmlElement | string)[];
}

/** Namespace URI to the canonical prefix; "" means the element carries no prefix. */
export type Namespaces = Record<string, string>;

const ENTITIES = new Map([["amp", "&"], ["lt", "<"], ["gt", ">"], ["quot", '"'], ["apos", "'"]]);
const XML_NS = "http://www.w3.org/XML/1998/namespace";
const SPACE = /[ \t\r\n]/;

const bad = (what: string): never => {
  throw new FormatError("INVALID_FILE", `the file is not valid XML: ${what}`);
};

const isXmlChar = (c: number): boolean =>
  c === 9 || c === 10 || c === 13 || (c >= 0x20 && c <= 0xd7ff) || (c >= 0xe000 && c <= 0xfffd) || (c >= 0x10000 && c <= 0x10ffff);

function decode(value: string): string {
  if (!value.includes("&")) return value;
  return value.replace(/&(#x[0-9a-f]+|#[0-9]+|[a-z]+);/gi, (_match, entity: string) => {
    if (entity[0] !== "#") return ENTITIES.get(entity) ?? bad(`the entity &${entity}; is not defined`);
    const code = entity[1] === "x" || entity[1] === "X" ? parseInt(entity.slice(2), 16) : parseInt(entity.slice(1), 10);
    return isXmlChar(code) ? String.fromCodePoint(code) : bad("a character reference is not allowed");
  });
}

/**
 * Read one XML part. A DOCTYPE is refused, so no entity and no external file is ever expanded.
 * Size, depth, and element count are capped. Prefixes are mapped to the canonical ones in
 * `namespaces`, so a file that binds `w` or `text` to another prefix still reads the same.
 */
export function parseXml(input: Uint8Array | string, limits: Limits, namespaces: Namespaces): XmlElement {
  if (typeof input !== "string" && input.length > limits.maxXmlBytes) throw new FormatError("LIMIT_EXCEEDED", `an XML part is larger than ${limits.maxXmlBytes} bytes`);
  const src = typeof input === "string" ? input : strFromU8(input);
  if (src.length > limits.maxXmlBytes) throw new FormatError("LIMIT_EXCEEDED", `an XML part is larger than ${limits.maxXmlBytes} bytes`);
  const encoding = /^﻿?<\?xml[^>]*\sencoding\s*=\s*["']([^"']+)["']/.exec(src)?.[1];
  if (encoding && !/^utf-?8$/i.test(encoding)) bad(`the encoding ${encoding} is not supported`);

  const root: XmlElement = { name: "", attrs: {}, children: [] };
  const stack: XmlElement[] = [root];
  const scopes: Record<string, string>[] = [{ xml: XML_NS }];
  let elements = 0;
  let i = 0;

  const qualify = (qname: string, scope: Record<string, string>, isElement: boolean): string => {
    const colon = qname.indexOf(":");
    const prefix = colon < 0 ? undefined : qname.slice(0, colon);
    const local = colon < 0 ? qname : qname.slice(colon + 1);
    if (prefix === undefined && !isElement) return local;
    const uri = scope[prefix ?? ""];
    const canonical = uri === undefined ? undefined : namespaces[uri];
    if (canonical === undefined) return qname;
    return canonical ? `${canonical}:${local}` : local;
  };

  while (i < src.length) {
    const top = stack[stack.length - 1];
    if (src[i] !== "<") {
      const end = src.indexOf("<", i);
      const text = src.slice(i, end < 0 ? src.length : end);
      if (stack.length > 1) top.children.push(decode(text));
      else if (text.trim() && text !== "﻿") bad("text outside the root element");
      i = end < 0 ? src.length : end;
      continue;
    }
    if (src.startsWith("<!--", i)) {
      const end = src.indexOf("-->", i + 4);
      if (end < 0) bad("a comment is not closed");
      i = end + 3;
    } else if (src.startsWith("<![CDATA[", i)) {
      const end = src.indexOf("]]>", i + 9);
      if (end < 0 || stack.length < 2) bad("a CDATA section is not valid here");
      top.children.push(src.slice(i + 9, end));
      i = end + 3;
    } else if (src.startsWith("<?", i)) {
      const end = src.indexOf("?>", i + 2);
      if (end < 0) bad("a processing instruction is not closed");
      i = end + 2;
    } else if (src.startsWith("<!", i)) {
      bad("a DOCTYPE is not allowed");
    } else if (src.startsWith("</", i)) {
      const end = src.indexOf(">", i);
      if (end < 0 || stack.length < 2) bad("an end tag has no start tag");
      const raw = src.slice(i + 2, end).trimEnd();
      const scope = scopes[scopes.length - 1];
      if (qualify(raw, scope, true) !== top.name) bad(`unexpected </${raw}>`);
      stack.pop();
      scopes.pop();
      i = end + 1;
    } else {
      // A start tag: name, attributes, then `>` or `/>`.
      let j = i + 1;
      while (j < src.length && !SPACE.test(src[j]) && src[j] !== ">" && src[j] !== "/") j++;
      const qname = src.slice(i + 1, j);
      if (!qname) bad("a stray <");
      const raw: [string, string][] = [];
      let selfClosing = false;
      for (;;) {
        while (j < src.length && SPACE.test(src[j])) j++;
        if (j >= src.length) bad("a tag is not closed");
        if (src[j] === ">") break;
        if (src[j] === "/" && src[j + 1] === ">") {
          selfClosing = true;
          j++;
          break;
        }
        const nameStart = j;
        while (j < src.length && !SPACE.test(src[j]) && src[j] !== "=" && src[j] !== ">" && src[j] !== "/") j++;
        const name = src.slice(nameStart, j);
        while (j < src.length && SPACE.test(src[j])) j++;
        if (!name || src[j] !== "=") bad(`the attribute "${name}" has no value`);
        j++;
        while (j < src.length && SPACE.test(src[j])) j++;
        const quote = src[j];
        if (quote !== '"' && quote !== "'") bad(`the attribute "${name}" is not quoted`);
        const close = src.indexOf(quote, j + 1);
        if (close < 0) bad("an attribute value is not closed");
        const value = src.slice(j + 1, close);
        if (value.includes("<")) bad("an attribute value holds a <");
        raw.push([name, decode(value)]);
        j = close + 1;
      }
      if (++elements > limits.maxXmlElements) throw new FormatError("LIMIT_EXCEEDED", `an XML part has more than ${limits.maxXmlElements} elements`);
      if (stack.length === 1 && top.children.some((c) => typeof c !== "string")) bad("it needs one root element");
      const scope = Object.create(scopes[scopes.length - 1]) as Record<string, string>;
      for (const [name, value] of raw) {
        if (name === "xmlns") scope[""] = value;
        else if (name.startsWith("xmlns:")) scope[name.slice(6)] = value;
      }
      const attrs: Record<string, string> = {};
      for (const [name, value] of raw) if (name !== "xmlns" && !name.startsWith("xmlns:")) attrs[qualify(name, scope, false)] = value;
      const element: XmlElement = { name: qualify(qname, scope, true), attrs, children: [] };
      top.children.push(element);
      if (!selfClosing) {
        if (stack.length > limits.maxXmlDepth) throw new FormatError("LIMIT_EXCEEDED", `XML nests deeper than ${limits.maxXmlDepth} levels`);
        stack.push(element);
        scopes.push(scope);
      }
      i = j + 1;
    }
  }
  const roots = root.children.filter((c): c is XmlElement => typeof c !== "string");
  if (stack.length !== 1 || roots.length !== 1) bad("it needs one root element and every tag closed");
  return roots[0];
}

export const elementsOf = (el: XmlElement | undefined): XmlElement[] => (el?.children ?? []).filter((c): c is XmlElement => typeof c !== "string");
export const childOf = (el: XmlElement | undefined, name: string): XmlElement | undefined => elementsOf(el).find((c) => c.name === name);

/** The text of an element and everything below it. */
export const textOf = (el: XmlElement | undefined): string => (el?.children ?? []).map((c) => (typeof c === "string" ? c : textOf(c))).join("");

/** Text safe for XML 1.0: drops the characters XML forbids, then escapes `& < > "`. */
export const xmlEscape = (value: unknown): string =>
  String(value ?? "")
    .replace(/[^\u0009\u000A\u000D -퟿-�\u{10000}-\u{10FFFF}]/gu, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/\r/g, "&#13;");

/**
 * A link target a file may carry: web, mail, in-page, or relative. Anything else, such as
 * `javascript:`, `file:`, or a network share (`//host/share`, `\\host\share`), returns undefined.
 * Writers report each dropped link as a warning.
 */
export const safeHref = (value: unknown): string | undefined => {
  const href = String(value ?? "").trim();
  if (/^[\\/]{2}/.test(href)) return undefined;
  return /^(https?:|mailto:|#|[^:]*$)/i.test(href) ? href : undefined;
};
