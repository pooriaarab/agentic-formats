/*
 * Direct tests for ooxml.ts. Ways it can fail, each with a case below:
 *  1 contentTypesXml / relsXml: well-formed, escaped, short and full type URIs, external mode
 *  2 resolveTarget: relative, absolute, `.`, `..`, past the root, unsafe names, empty
 *  3 readOoxmlPackage: round trip, strict namespaces, no root rel, external root rel, wrong content
 *    type, missing main part, rels that leave the package, external rels kept, rels without id/type,
 *    broken XML, size caps
 */
import { strToU8, zipSync } from "fflate";
import { describe, expect, test } from "vitest";
import { DEFAULT_LIMITS, FormatError, withLimits } from "./limits.js";
import { contentTypesXml, OOXML_NAMESPACES, readOoxmlPackage, relsXml, resolveTarget } from "./ooxml.js";
import { elementsOf, parseXml } from "./xml.js";

const MAIN = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const code = (fn: () => unknown) => {
  try {
    fn();
  } catch (e) {
    if (e instanceof FormatError) return e.code;
    throw e;
  }
  return "none";
};
const zip = (files: Record<string, string>) => zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));
const types = (part = "word/document.xml", type = MAIN) => contentTypesXml({ xml: "application/xml" }, { [part]: type });
const rootRels = (target = "word/document.xml", extra = "") =>
  relsXml([{ id: "rId1", type: "officeDocument", target }]).replace("</Relationships>", `${extra}</Relationships>`);
const pkg = (extra: Record<string, string> = {}, over: Record<string, string> = {}) =>
  zip({ "[Content_Types].xml": types(), "_rels/.rels": rootRels(), "word/document.xml": "<w:document/>", ...extra, ...over });
const read = (b: Uint8Array, limits = DEFAULT_LIMITS) => readOoxmlPackage(b, limits, "docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main");

describe("contentTypesXml and relsXml", () => {
  test("content types parse back", () => {
    const root = parseXml(contentTypesXml({ png: "image/png" }, { "word/document.xml": MAIN }), DEFAULT_LIMITS, OOXML_NAMESPACES);
    expect(elementsOf(root).map((e) => [e.name, e.attrs.Extension ?? e.attrs.PartName])).toEqual([["Default", "png"], ["Override", "/word/document.xml"]]);
  });
  test("rels use the short type and the full type", () => {
    const x = relsXml([{ id: "a", type: "image", target: "media/a.png" }, { id: "b", type: "http://x/y/link", target: "u" }]);
    expect(x).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"');
    expect(x).toContain('Type="http://x/y/link"');
  });
  test("rels escape targets and mark external ones", () => {
    const x = relsXml([{ id: "a", type: "hyperlink", target: 'http://e.com/?a=1&b="2"', external: true }]);
    expect(x).toContain('Target="http://e.com/?a=1&amp;b=&quot;2&quot;" TargetMode="External"');
    expect(parseXml(x, DEFAULT_LIMITS, OOXML_NAMESPACES).children).toHaveLength(1);
  });
  test("empty lists are well formed", () => {
    expect(() => parseXml(relsXml([]), DEFAULT_LIMITS, OOXML_NAMESPACES)).not.toThrow();
    expect(() => parseXml(contentTypesXml({}, {}), DEFAULT_LIMITS, OOXML_NAMESPACES)).not.toThrow();
  });
});

describe("resolveTarget", () => {
  test.each([
    ["word/document.xml", "media/a.png", "word/media/a.png"],
    ["word/document.xml", "../customXml/a.xml", "customXml/a.xml"],
    ["word/document.xml", "/docProps/core.xml", "docProps/core.xml"],
    ["word/document.xml", "./a.xml", "word/a.xml"],
    ["word/document.xml", "a//b.xml", "word/a/b.xml"],
    ["", "word/document.xml", "word/document.xml"],
    ["xl/worksheets/sheet1.xml", "../../a.xml", "a.xml"],
  ])("%s + %s = %s", (source, target, want) => expect(resolveTarget(source, target)).toBe(want));
  test.each([
    ["word/document.xml", "../../etc/passwd"],
    ["word/document.xml", "/../a.xml"],
    ["document.xml", "../a.xml"],
    ["word/document.xml", "a\\b.xml"],
    ["word/document.xml", "media/"],
    ["word/document.xml", "."],
    ["word/document.xml", ".."],
    ["word/document.xml", ""],
    ["word/document.xml", "/"],
    ["word/document.xml", "a\u0000b.xml"],
  ])("%s + %j leaves the package", (source, target) => expect(resolveTarget(source, target)).toBeUndefined());
});

describe("readOoxmlPackage", () => {
  test("finds the main part and parses it", () => {
    const file = read(pkg());
    expect(file.mainPart).toBe("word/document.xml");
    expect(file.parse("word/document.xml").name).toBe("w:document");
  });
  test("reads strict namespaces to the same prefix", () => {
    const strict = '<document xmlns="http://purl.oclc.org/ooxml/wordprocessingml/main"/>';
    expect(read(pkg({ "word/document.xml": strict })).parse("word/document.xml").name).toBe("w:document");
  });
  test("parse throws for a part that is not there", () => expect(code(() => read(pkg()).parse("word/nope.xml"))).toBe("INVALID_FILE"));
  test("reads part rels, caches them, and keeps external ones", () => {
    const rels = relsXml([
      { id: "rId1", type: "image", target: "media/a.png" },
      { id: "rId2", type: "hyperlink", target: "http://e.com/x", external: true },
      { id: "rId3", type: "image", target: "../../../escape.png" },
    ]);
    const file = read(pkg({ "word/_rels/document.xml.rels": rels, "word/media/a.png": "png" }));
    const map = file.rels("word/document.xml");
    expect(map.get("rId1")).toEqual({ id: "rId1", kind: "image", target: "word/media/a.png", external: false });
    expect(map.get("rId2")).toEqual({ id: "rId2", kind: "hyperlink", target: "http://e.com/x", external: true });
    expect(map.has("rId3")).toBe(false);
    expect(file.rels("word/document.xml")).toBe(map);
  });
  test("a part with no rels has an empty map", () => expect(read(pkg()).rels("word/other.xml").size).toBe(0));
  test("skips rels without an id or a type", () => {
    const x = relsXml([]).replace("</Relationships>", '<Relationship Type="t/image" Target="a"/><Relationship Id="z" Target="a"/></Relationships>');
    expect(read(pkg({ "word/_rels/document.xml.rels": x })).rels("word/document.xml").size).toBe(0);
  });
  test("rejects no root rels", () => expect(code(() => read(zip({ "[Content_Types].xml": types(), "word/document.xml": "<a/>" })))).toBe("INVALID_FILE"));
  test("rejects an external officeDocument rel", () => {
    const rels = relsXml([{ id: "rId1", type: "officeDocument", target: "http://e.com/d.xml", external: true }]);
    expect(code(() => read(pkg({}, { "_rels/.rels": rels })))).toBe("INVALID_FILE");
  });
  test("rejects a root rel that leaves the package", () => expect(code(() => read(pkg({}, { "_rels/.rels": rootRels("../../a.xml") })))).toBe("INVALID_FILE"));
  test("rejects a content type of another format", () => {
    const t = types("word/document.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml");
    expect(code(() => read(pkg({}, { "[Content_Types].xml": t })))).toBe("INVALID_FILE");
  });
  test("rejects a missing content type for the main part", () => {
    expect(code(() => read(pkg({}, { "[Content_Types].xml": types("word/other.xml") })))).toBe("INVALID_FILE");
  });
  test("rejects missing [Content_Types].xml", () => {
    const files = { "_rels/.rels": rootRels(), "word/document.xml": "<a/>" };
    expect(code(() => read(zip(files)))).toBe("INVALID_FILE");
  });
  test("rejects broken rels XML", () => expect(code(() => read(pkg({}, { "_rels/.rels": "<Relationships>" })))).toBe("INVALID_FILE"));
  test("rejects bytes that are not a zip", () => expect(code(() => read(strToU8("nope")))).toBe("INVALID_FILE"));
  test("drops entries of other types", () => expect(Object.keys(read(pkg({ "word/macro.bin": "x", "a.exe": "x" })).files)).not.toContain("word/macro.bin"));
  test("keeps images", () => expect(Object.keys(read(pkg({ "word/media/a.PNG": "x", "word/media/b.jpeg": "x" })).files)).toEqual(expect.arrayContaining(["word/media/a.PNG", "word/media/b.jpeg"])));
  test("enforces the XML size cap lazily, when the part is parsed", () => {
    const b = pkg({ "word/document.xml": `<w:document>${"x".repeat(2000)}</w:document>` });
    const small = withLimits({ maxXmlBytes: 1000 });
    expect(code(() => read(b, small).parse("word/document.xml"))).toBe("LIMIT_EXCEEDED");
  });
  test("enforces the entry cap", () => expect(code(() => read(pkg(), withLimits({ maxEntries: 2 })))).toBe("LIMIT_EXCEEDED"));
  test("enforces the input cap", () => expect(code(() => read(pkg(), withLimits({ maxInputBytes: 10 })))).toBe("LIMIT_EXCEEDED"));
});
