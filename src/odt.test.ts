/*
 * ODT, end to end: the real writer, the real reader, an independent XML parser, and LibreOffice.
 * The ways it can fail are written first; each has a case below.
 *
 *  1 text is lost or changed: marks, links, runs of spaces, tabs, line breaks
 *  2 structure is lost: headings, nested lists, ordered or not, tables with a header and empty cells
 *  3 an image or a footnote is lost
 *  4 the package is malformed: an XML part does not parse, `mimetype` is not first and stored,
 *    the manifest root carries the wrong version attribute
 *  5 the same model gives different bytes
 *  6 a size limit is missing: an image or the whole export grows without a bound
 *  7 an image frame takes its size from a hostile header or model
 *  8 an unsafe link or a bad image drops without a word
 *  9 a hostile file is read: a DOCTYPE, a wrong media type, a remote image, a script link
 * 10 LibreOffice cannot open our file, or its re-save does not read back to the same model
 */
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, expect, test } from "vitest";
import { FormatError } from "./limits.js";
import { normalizeTextDoc, type TextDoc } from "./model.js";
import { readOdt, writeOdt } from "./odt.js";
import { LibreOffice, hasLibreOffice, libreOfficeRequired, sha256, writeArtifact, type RoundTripRecord } from "./testing/libreoffice.js";
import { writeZip } from "./zip.js";
import { withLimits } from "./limits.js";

const PNG = Uint8Array.from(atob("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="), (c) => c.charCodeAt(0));

const fixture: TextDoc = {
  title: "Quarterly notes & <plans>",
  blocks: [
    { type: "heading", level: 1, content: [{ type: "text", text: "Plan" }] },
    { type: "heading", level: 3, content: [{ type: "text", text: "Details" }] },
    {
      type: "paragraph",
      content: [
        { type: "text", text: "Plain, " },
        { type: "text", text: "bold", marks: ["bold"] },
        { type: "text", text: " " },
        { type: "text", text: "italic and underline", marks: ["italic", "underline"] },
        { type: "text", text: " " },
        { type: "text", text: "struck", marks: ["strike"] },
        { type: "text", text: " " },
        { type: "text", text: "code()", marks: ["code"] },
        { type: "text", text: " x" },
        { type: "text", text: "2", marks: ["superscript"] },
        { type: "text", text: " H" },
        { type: "text", text: "2", marks: ["subscript"] },
        { type: "text", text: "O, " },
        { type: "text", text: "a link", href: "https://example.org/a?b=1&c=2" },
        { type: "text", text: " and " },
        { type: "text", text: "bold link", marks: ["bold"], href: "mailto:team@example.org" },
        { type: "text", text: "." },
      ],
    },
    { type: "paragraph", content: [{ type: "text", text: "  lead, a    gap,\ttab" }, { type: "break" }, { type: "text", text: "after break  " }] },
    { type: "paragraph", content: [{ type: "text", text: "Entities: 5 < 6 & \"quoted\" 'single' ünï ✓ 𝒳" }] },
    {
      type: "list",
      ordered: false,
      items: [
        { content: [{ type: "text", text: "one" }] },
        { content: [{ type: "text", text: "two" }], sublist: { type: "list", ordered: true, items: [{ content: [{ type: "text", text: "two-a" }] }, { content: [{ type: "text", text: "two-b" }], sublist: { type: "list", ordered: false, items: [{ content: [{ type: "text", text: "deep" }] }] } }] } },
      ],
    },
    { type: "list", ordered: true, items: [{ content: [{ type: "text", text: "first" }] }, { content: [{ type: "text", text: "second", marks: ["bold"] }] }] },
    {
      type: "table",
      header: true,
      rows: [
        [[{ type: "paragraph", content: [{ type: "text", text: "Name" }] }], [{ type: "paragraph", content: [{ type: "text", text: "Qty" }] }]],
        [[{ type: "paragraph", content: [{ type: "text", text: "Apples" }] }], []],
        [[{ type: "paragraph", content: [{ type: "text", text: "Two" }] }, { type: "paragraph", content: [{ type: "text", text: "paragraphs" }] }], [{ type: "paragraph", content: [{ type: "text", text: "7", marks: ["bold"] }] }]],
      ],
    },
    { type: "paragraph", content: [{ type: "text", text: "Image: " }, { type: "image", data: PNG, mime: "image/png", alt: "a red dot", width: 72, height: 36 }, { type: "text", text: " ends." }] },
    {
      type: "paragraph",
      content: [{ type: "text", text: "Claim" }, { type: "footnote", content: [{ type: "text", text: "See " }, { type: "text", text: "source", marks: ["italic"] }, { type: "break" }, { type: "text", text: "page 4." }] }, { type: "text", text: " holds." }],
    },
  ],
};

const parts = (bytes: Uint8Array) => unzipSync(bytes);
const wellFormed = (xml: string): void => {
  const seen: string[] = [];
  new DOMParser({ onError: (level, message) => seen.push(`${level}: ${message}`) }).parseFromString(xml, "text/xml");
  expect(seen).toEqual([]);
};

describe("write then read", () => {
  const written = writeOdt(fixture);

  test("1-3 the model reads back unchanged, with no warnings", () => {
    const read = readOdt(written.bytes);
    expect(read.model).toEqual(normalizeTextDoc(fixture));
    expect(read.warnings).toEqual([]);
    expect(written.warnings).toEqual([]);
  });

  test("4 every XML part parses with an independent parser", () => {
    const files = parts(written.bytes);
    const xml = Object.keys(files).filter((n) => n.endsWith(".xml"));
    expect(xml.sort()).toEqual(["META-INF/manifest.xml", "content.xml", "meta.xml", "styles.xml"]);
    for (const name of xml) wellFormed(strFromU8(files[name]));
  });

  test("4 `mimetype` is first and stored, and the manifest carries `manifest:version`", () => {
    const bytes = written.bytes;
    expect(strFromU8(bytes.subarray(30, 38))).toBe("mimetype");
    expect(new DataView(bytes.buffer, bytes.byteOffset).getUint16(8, true)).toBe(0); // method: stored
    const manifest = strFromU8(parts(bytes)["META-INF/manifest.xml"]);
    expect(manifest).toMatch(/<manifest:manifest [^>]*manifest:version="1\.3"/);
    expect(manifest).not.toMatch(/<manifest:manifest [^>]*office:version/);
    expect(manifest).toContain('manifest:full-path="Pictures/Image1.png"');
  });

  test("5 the same model gives the same bytes", () => {
    expect(writeOdt(fixture).bytes).toEqual(written.bytes);
  });
});

describe("write guards", () => {
  const withImage = (data: Uint8Array, extra: object = {}): TextDoc => ({ blocks: [{ type: "paragraph", content: [{ type: "image", data, mime: "image/png", ...extra }] }] });
  const frameOf = (doc: TextDoc) => {
    const content = strFromU8(parts(writeOdt(doc).bytes)["content.xml"]);
    return { w: Number(/svg:width="([0-9.]+)cm"/.exec(content)![1]) * (72 / 2.54), h: Number(/svg:height="([0-9.]+)cm"/.exec(content)![1]) * (72 / 2.54) };
  };

  test("6 an image over the entry limit, and an export over the total limit, are refused", () => {
    const big = new Uint8Array(PNG.length + 4000);
    big.set(PNG);
    expect(() => writeOdt(withImage(big), { maxEntryBytes: 1000 })).toThrow(FormatError);
    expect(() => writeOdt(withImage(PNG), { maxTotalBytes: 3000 })).toThrow(/export is larger/);
  });

  test("7 the frame size comes from a usable model size, else the header, and always fits 22 inches", () => {
    const png = (w: number, h: number) => {
      const copy = PNG.slice();
      new DataView(copy.buffer).setUint32(16, w);
      new DataView(copy.buffer).setUint32(20, h);
      return copy;
    };
    const hostile = frameOf(withImage(png(4_000_000_000, 4_000_000_000)));
    expect(hostile.w).toBeLessThanOrEqual(1584.1);
    expect(hostile.h).toBeLessThanOrEqual(1584.1);
    const zero = frameOf(withImage(png(0, 0)));
    expect([Math.round(zero.w), Math.round(zero.h)]).toEqual([240, 180]);
    const header = frameOf(withImage(png(200, 100)));
    expect([Math.round(header.w), Math.round(header.h)]).toEqual([150, 75]);
    for (const bad of [Number.NaN, -5, 0, Infinity]) {
      const f = frameOf(withImage(png(200, 100), { width: bad, height: bad }));
      expect([Math.round(f.w), Math.round(f.h)]).toEqual([150, 75]);
    }
    const huge = frameOf(withImage(PNG, { width: 1e12, height: 5e11 }));
    expect(huge.w).toBeLessThanOrEqual(1584.1);
    expect(huge.w / huge.h).toBeCloseTo(2, 1);
  });

  test("8 an unsafe link and a bad image are reported", () => {
    const doc: TextDoc = {
      blocks: [{ type: "paragraph", content: [{ type: "text", text: "click", href: "javascript:alert(1)" }, { type: "image", data: new TextEncoder().encode("<svg onload=alert(1)/>"), mime: "image/png" }] }],
    };
    const { bytes, warnings } = writeOdt(doc);
    expect(warnings).toEqual(["a link with an unsafe target was written as plain text", "an image is not a PNG, JPEG, or GIF: not written"]);
    const content = strFromU8(parts(bytes)["content.xml"]);
    expect(content).toContain("click");
    expect(content).not.toMatch(/javascript|<svg|Pictures/);
  });
});

describe("hostile files", () => {
  const odt = (content: string, mime = "application/vnd.oasis.opendocument.text", extra: { name: string; data: Uint8Array }[] = []) =>
    writeZip([{ name: "mimetype", data: mime, store: true }, { name: "content.xml", data: content }, ...extra], withLimits());
  const wrap = (body: string) =>
    `<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text>${body}</office:text></office:body></office:document-content>`;

  test("9 a DOCTYPE with an entity is refused", () => {
    const xml = `<?xml version="1.0"?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]>` + wrap("<text:p>&e;</text:p>").split("?>")[1];
    expect(() => readOdt(odt(xml))).toThrow(/DOCTYPE/);
  });

  test("9 another media type, a zip with no content, and a non-zip are refused", () => {
    expect(() => readOdt(odt(wrap(""), "application/vnd.oasis.opendocument.spreadsheet"))).toThrow(/not an ODT/);
    expect(() => readOdt(writeZip([{ name: "mimetype", data: "application/vnd.oasis.opendocument.text", store: true }], withLimits()))).toThrow(/no content\.xml/);
    expect(() => readOdt(new TextEncoder().encode("hello"))).toThrow(FormatError);
  });

  test("9 a remote image is not loaded, a script link reads as text, a text with a template media type reads", () => {
    const xml = wrap(
      `<text:p>see <text:a xlink:href="javascript:alert(1)">this</text:a><draw:frame><draw:image xlink:href="https://evil.example/pixel.png"/></draw:frame></text:p>`,
    );
    const read = readOdt(odt(xml, "application/vnd.oasis.opendocument.text-template"));
    expect(read.model.blocks).toEqual([{ type: "paragraph", content: [{ type: "text", text: "see this" }] }]);
    expect(read.warnings.sort()).toEqual(["a link with an unsafe target was read as plain text", "a remote image: not loaded"]);
  });

  test("9 a table that repeats a cell a million times is refused", () => {
    const xml = wrap(`<table:table xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><table:table-row><table:table-cell table:number-columns-repeated="1000000"/></table:table-row></table:table>`);
    expect(() => readOdt(odt(xml))).toThrow(/repeats/);
  });
});

describe.skipIf(!hasLibreOffice && !libreOfficeRequired)("LibreOffice round trip", () => {
  const lo = hasLibreOffice ? new LibreOffice() : undefined;
  const records: RoundTripRecord[] = [];
  afterAll(() => {
    lo?.dispose();
    writeArtifact("odt-roundtrip", records);
  });

  test("10 LibreOffice opens our file, prints it, and its re-save reads back to the same model", { timeout: 300_000 }, () => {
    expect(hasLibreOffice).toBe(true);
    const { bytes } = writeOdt(fixture);
    const pdf = lo!.convert(bytes, "odt", "pdf");
    expect(strFromU8(pdf.subarray(0, 5))).toBe("%PDF-");
    const text = strFromU8(lo!.convert(bytes, "odt", "txt:Text"));
    for (const word of ["Plan", "Details", "two-b", "deep", "Apples", "paragraphs", "ends.", "Claim", "holds."]) expect(text).toContain(word);
    const resaved = lo!.convert(bytes, "odt", "odt");
    const back = readOdt(resaved);
    const equal = JSON.stringify(back.model) === JSON.stringify(normalizeTextDoc(fixture));
    records.push({ format: "odt", case: "fixture", written: sha256(bytes), resaved: sha256(resaved), pdfBytes: pdf.length, equal });
    expect(back.model).toEqual(normalizeTextDoc(fixture));
  });
});
