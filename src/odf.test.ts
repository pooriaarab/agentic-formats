/*
 * Direct tests for odf.ts. Ways it can fail, each with a case below:
 *  1 spaced: tab, newline, leading, trailing, inner runs of spaces, markup characters
 *  2 spacedBound: never lower than the real length of spaced()
 *  3 writeOdfPackage: mimetype first and stored, manifest lists images, title escaped, size caps
 *  4 readOdfPackage: round trip, wrong mime, no mimetype, no content.xml, bad XML, no title,
 *    parts that the reader does not want are dropped
 */
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, test } from "vitest";
import { DEFAULT_LIMITS, FormatError, withLimits } from "./limits.js";
import { ODF_NAMESPACES, ODF_NS_ATTRS, readOdfPackage, spaced, spacedBound, writeOdfPackage, type OdfParts } from "./odf.js";

const ODT = "application/vnd.oasis.opendocument.text";
const root = (body: string) => `<?xml version="1.0"?><office:document-content ${ODF_NS_ATTRS}>${body}</office:document-content>`;
const parts = (over: Partial<OdfParts> = {}): OdfParts => ({
  mime: ODT,
  content: root("<office:body><office:text><text:p>Hi</text:p></office:text></office:body>"),
  styles: root(""),
  images: [],
  ...over,
});
const code = (fn: () => unknown) => {
  try {
    fn();
  } catch (e) {
    if (e instanceof FormatError) return e.code;
    throw e;
  }
  return "none";
};

describe("spaced", () => {
  test("keeps plain text and single inner spaces", () => expect(spaced("a b c")).toBe("a b c"));
  test("turns tab and newline into elements", () => expect(spaced("a\tb\nc")).toBe("a<text:tab/>b<text:line-break/>c"));
  test("keeps one space of an inner run and counts the rest", () => expect(spaced("a   b")).toBe('a <text:s text:c="2"/>b'));
  test("counts every leading and trailing space", () => expect(spaced("  a ")).toBe('<text:s text:c="2"/>a<text:s text:c="1"/>'));
  test("a lone space is an element", () => expect(spaced(" ")).toBe('<text:s text:c="1"/>'));
  test("escapes markup", () => expect(spaced("<a & b>")).toBe("&lt;a &amp; b&gt;"));
  test("empty text stays empty", () => expect(spaced("")).toBe(""));
});

describe("spacedBound", () => {
  const samples = ["", " ", "  ", "a b", "a  b", " a ", "\n\n\n", "\t\t", "<<<&&&>>>", '"""', "\r\n", "a   b   c   ", "x".repeat(50) + "\n" + " ".repeat(30)];
  test.each(samples)("is at least the length of spaced(%j)", (s) => expect(spacedBound(s)).toBeGreaterThanOrEqual(spaced(s).length));
  test("is the plain length for text with nothing to expand", () => expect(spacedBound("abc def")).toBe(7));
});

describe("writeOdfPackage", () => {
  const png = new Uint8Array([137, 80, 78, 71]);
  const bytes = writeOdfPackage(parts({ title: "A & B", images: [{ path: "Pictures/i.png", mime: "image/png", data: png }] }), DEFAULT_LIMITS);

  test("puts mimetype first, stored, with no extra field", () => {
    expect(strFromU8(bytes.subarray(30, 38))).toBe("mimetype");
    expect(new DataView(bytes.buffer, bytes.byteOffset).getUint16(8, true)).toBe(0);
    expect(strFromU8(bytes.subarray(38, 38 + ODT.length))).toBe(ODT);
  });
  test("writes every part and the image", () => {
    const files = unzipSync(bytes);
    expect(Object.keys(files).sort()).toEqual(["META-INF/manifest.xml", "Pictures/i.png", "content.xml", "meta.xml", "mimetype", "styles.xml"]);
    expect([...files["Pictures/i.png"]]).toEqual([...png]);
  });
  test("lists the image in the manifest", () => {
    const manifest = strFromU8(unzipSync(bytes)["META-INF/manifest.xml"]);
    expect(manifest).toContain('manifest:full-path="Pictures/i.png" manifest:media-type="image/png"');
    expect(manifest).toContain(`manifest:media-type="${ODT}"`);
  });
  test("escapes the title", () => expect(strFromU8(unzipSync(bytes)["meta.xml"])).toContain("<dc:title>A &amp; B</dc:title>"));
  test("writes no title element without a title", () => {
    expect(strFromU8(unzipSync(writeOdfPackage(parts(), DEFAULT_LIMITS))["meta.xml"])).not.toContain("dc:title");
  });
  test("rejects output past the size cap", () => {
    expect(code(() => writeOdfPackage(parts({ content: "x".repeat(5000) }), withLimits({ maxTotalBytes: 1000 })))).toBe("LIMIT_EXCEEDED");
  });
});

describe("readOdfPackage", () => {
  const read = (b: Uint8Array, limits = DEFAULT_LIMITS, mimes = [ODT]) => readOdfPackage(b, limits, "odt", mimes);
  const zip = (files: Record<string, string>) => zipSync(Object.fromEntries(Object.entries(files).map(([k, v]) => [k, strToU8(v)])));

  test("round trips content, styles, title, and images", () => {
    const data = new Uint8Array([1, 2, 3]);
    const file = read(writeOdfPackage(parts({ title: "Plan", images: [{ path: "Pictures/a.png", mime: "image/png", data }] }), DEFAULT_LIMITS));
    expect(file.title).toBe("Plan");
    expect(file.content.name).toBe("office:document-content");
    expect(file.styles?.name).toBe("office:document-content");
    expect([...file.files["Pictures/a.png"]]).toEqual([1, 2, 3]);
  });
  test("has no title when meta.xml is absent or the title is blank", () => {
    expect(read(zip({ mimetype: ODT, "content.xml": root("") })).title).toBeUndefined();
    expect("title" in read(writeOdfPackage(parts({ title: "  " }), DEFAULT_LIMITS))).toBe(false);
  });
  test("trims the mimetype", () => expect(read(zip({ mimetype: `${ODT}\n`, "content.xml": root("") })).content.name).toBe("office:document-content"));
  test("accepts any listed mime", () => expect(read(zip({ mimetype: "x/y", "content.xml": root("") }), DEFAULT_LIMITS, [ODT, "x/y"]).files.mimetype).toBeDefined());
  test("rejects another mime", () => expect(code(() => read(zip({ mimetype: "application/zip", "content.xml": root("") })))).toBe("INVALID_FILE"));
  test("rejects a missing mimetype", () => expect(code(() => read(zip({ "content.xml": root("") })))).toBe("INVALID_FILE"));
  test("rejects a missing content.xml", () => expect(code(() => read(zip({ mimetype: ODT })))).toBe("INVALID_FILE"));
  test("rejects broken content XML", () => expect(code(() => read(zip({ mimetype: ODT, "content.xml": "<a><b></a>" })))).toBe("INVALID_FILE"));
  test("rejects broken meta XML", () => expect(code(() => read(zip({ mimetype: ODT, "content.xml": root(""), "meta.xml": "<a>" })))).toBe("INVALID_FILE"));
  test("rejects bytes that are not a zip", () => expect(code(() => read(strToU8("not a zip")))).toBe("INVALID_FILE"));
  test("drops parts it does not read", () => {
    const file = read(zip({ mimetype: ODT, "content.xml": root(""), "Scripts/a.js": "x", "Pictures/": "", "settings.xml": "<a/>" }));
    expect(Object.keys(file.files).sort()).toEqual(["content.xml", "mimetype"]);
  });
  test("enforces the XML size cap", () => {
    expect(code(() => read(zip({ mimetype: ODT, "content.xml": root("<office:body>" + "x".repeat(2000) + "</office:body>") }), withLimits({ maxXmlBytes: 500 })))).toBe("LIMIT_EXCEEDED");
  });
  test("enforces the input size cap", () => {
    expect(code(() => read(writeOdfPackage(parts(), DEFAULT_LIMITS), withLimits({ maxInputBytes: 10 })))).toBe("LIMIT_EXCEEDED");
  });
  test("maps the namespaces the writer declares to the spec prefixes", () => {
    for (const m of ODF_NS_ATTRS.matchAll(/xmlns:(\w+)="([^"]+)"/g)) expect(ODF_NAMESPACES[m[2]]).toBe(m[1]);
  });
});
