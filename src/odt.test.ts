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
 * 11 a table makes billions of cells from a few repeat counts (rows x cells x repeats, over many tables)
 * 12 a style parent chain is long, loops, or names a key of Object.prototype
 * 13 one big paragraph in a repeated table cell fills the output once for each copy
 * 14 a few bytes of <text:s> expand to millions of spaces; nested spans stack the same marks again
 *    and again; one image part named by many frames; a writer holds one object many times
 */
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, expect, test } from "vitest";
import { FormatError } from "./limits.js";
import { normalizeTextDoc, type Block, type TextDoc } from "./model.js";
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

  const TABLE_NS = `xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"`;
  const cells = (rowRepeat: number, cellRepeat: number, rows = 1) =>
    `<table:table ${TABLE_NS}>${`<table:table-row table:number-rows-repeated="${rowRepeat}"><table:table-cell table:number-columns-repeated="${cellRepeat}"/></table:table-row>`.repeat(rows)}</table:table>`;
  const limitExceeded = (run: () => unknown) => {
    try {
      run();
    } catch (error) {
      expect(error).toBeInstanceOf(FormatError);
      expect((error as FormatError).code).toBe("LIMIT_EXCEEDED");
      return;
    }
    throw new Error("expected LIMIT_EXCEEDED");
  };

  test("11 a thousand rows that repeat a thousand times over a thousand cells are refused fast", () => {
    const started = Date.now();
    limitExceeded(() => readOdt(odt(wrap(cells(1000, 1000, 1000)))));
    expect(Date.now() - started).toBeLessThan(2000);
  });

  test("11 rows x cells x repeats count against one budget for the whole document, over many tables", () => {
    const budget = withLimits({ maxCells: 100 });
    expect(readOdt(odt(wrap(cells(10, 10))), budget).model.blocks).toHaveLength(1);
    limitExceeded(() => readOdt(odt(wrap(cells(10, 11))), budget));
    limitExceeded(() => readOdt(odt(wrap(cells(1, 1, 101))), budget));
    limitExceeded(() => readOdt(odt(wrap(cells(1, 60) + cells(1, 41))), budget));
  });

  test("11 a nested table counts against the same budget", () => {
    const inner = `<table:table ${TABLE_NS}><table:table-row table:number-rows-repeated="10"><table:table-cell table:number-columns-repeated="10"/></table:table-row></table:table>`;
    const nested = `<table:table ${TABLE_NS}><table:table-row table:number-rows-repeated="10"><table:table-cell table:number-columns-repeated="10">${inner}</table:table-cell></table:table-row></table:table>`;
    // 100 cells in the inner table, then 10 x 10 copies of it: 100 + 10,000. A budget of 300 rejects only a reader that multiplies by the outer repeats.
    limitExceeded(() => readOdt(odt(wrap(nested)), withLimits({ maxCells: 300 })));
    expect(readOdt(odt(wrap(nested)), withLimits({ maxCells: 10_100 })).model.blocks).toHaveLength(1);
  });

  test("13 a big paragraph in a repeated cell or row counts once for each copy", () => {
    const caps = withLimits({ maxTextChars: 60_000 });
    const cell = (attrs: string) => `<table:table-cell ${attrs}><text:p>${"x".repeat(50_000)}</text:p></table:table-cell>`;
    const row = (cellAttrs: string, rowAttrs = "") => `<table:table ${TABLE_NS}><table:table-row ${rowAttrs}>${cell(cellAttrs)}</table:table-row></table:table>`;
    expect(readOdt(odt(wrap(row('table:number-columns-repeated="2"'))), caps).model.blocks).toHaveLength(1);
    limitExceeded(() => readOdt(odt(wrap(row('table:number-columns-repeated="3"'))), caps));
    limitExceeded(() => readOdt(odt(wrap(row("", 'table:number-rows-repeated="3"'))), caps));
    limitExceeded(() => readOdt(odt(wrap(row('table:number-columns-repeated="1000"', 'table:number-rows-repeated="1000"'))), caps));
  });

  test("14 text:s runs count against the text cap, and a long run is shortened with a warning", () => {
    const spaces = (n: number, c: number) => `<text:p>a${`<text:s text:c="${c}"/>`.repeat(n)}b</text:p>`;
    const caps = withLimits({ maxTextChars: 10_000 });
    expect(readOdt(odt(wrap(spaces(9, 1000))), caps).model.blocks).toHaveLength(1);
    limitExceeded(() => readOdt(odt(wrap(spaces(11, 1000))), caps));
    const shortened = readOdt(odt(wrap(spaces(1, 5000))));
    expect(shortened.warnings).toEqual(["a run of more than 1000 spaces was shortened"]);
    const started = Date.now();
    limitExceeded(() => readOdt(odt(wrap(spaces(200_000, 1000)))));
    expect(Date.now() - started).toBeLessThan(5000);
  });

  test("14 nested spans with a long style chain keep seven marks at most", () => {
    const marks = `<style:text-properties fo:font-weight="bold" fo:font-style="italic" style:text-underline-style="solid" style:text-line-through-style="solid" fo:font-family="monospace" style:text-position="super 58%" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"/>`;
    const styles = Array.from({ length: 50 }, (_, i) => `<style:style style:name="m${i}"${i ? ` style:parent-style-name="m${i - 1}"` : ""}>${marks}</style:style>`).join("");
    const spans = `<text:span text:style-name="m49">`.repeat(60) + "x" + "</text:span>".repeat(60);
    const started = Date.now();
    const read = readOdt(odt(withStyles(styles, `<text:p>${spans}</text:p>`)));
    expect(Date.now() - started).toBeLessThan(2000);
    expect(read.model.blocks).toEqual([{ type: "paragraph", content: [{ type: "text", text: "x", marks: ["bold", "italic", "underline", "strike", "code", "superscript"] }] }]);
  });

  test("14 one image part named by many frames counts its bytes for each frame", () => {
    const frame = `<draw:frame><draw:image xlink:href="Pictures/a.png"/></draw:frame>`;
    const body = wrap(`<text:p>${frame.repeat(5)}</text:p>`);
    const padded = new Uint8Array(10_000);
    padded.set(PNG);
    const picture = [{ name: "Pictures/a.png", data: padded }];
    // The zip holds the image once (about 11,000 bytes); five frames name it, so the reader counts 50,000.
    expect(readOdt(odt(body, "application/vnd.oasis.opendocument.text", picture), withLimits({ maxTotalBytes: 60_000 })).model.blocks).toHaveLength(1);
    limitExceeded(() => readOdt(odt(body, "application/vnd.oasis.opendocument.text", picture), withLimits({ maxTotalBytes: 40_000 })));
  });

  test("14 a writer stops at the byte cap when one object stands in many places", () => {
    const row = Array.from({ length: 1000 }, () => [{ type: "paragraph" as const, content: [{ type: "text" as const, text: "x" }] }]);
    const table = { type: "table" as const, rows: Array.from({ length: 2000 }, () => row) };
    const started = Date.now();
    limitExceeded(() => writeOdt({ blocks: [table] }, withLimits({ maxTotalBytes: 5_000_000 })));
    expect(Date.now() - started).toBeLessThan(5000);
    // Forty objects that stand for 2^40 tables: each level holds the level below twice.
    const nest = (k: number): Block => {
      if (k === 0) return { type: "paragraph", content: [{ type: "text", text: "x" }] };
      const inner = nest(k - 1);
      return { type: "table", rows: [[[inner], [inner]]] };
    };
    limitExceeded(() => writeOdt({ blocks: [nest(40)] }, withLimits({ maxTotalBytes: 1_000_000 })));
  });

  const STYLE_NS = `xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"`;
  const withStyles = (styles: string, body: string) =>
    wrap(body).replace("<office:body>", `<office:automatic-styles ${STYLE_NS}>${styles}</office:automatic-styles><office:body>`);
  const chain = (length: number) =>
    Array.from({ length }, (_, i) => `<style:style style:name="s${i}"${i ? ` style:parent-style-name="s${i - 1}"` : ""}/>`).join("");

  test("12 a style chain of twenty thousand styles is refused fast, however many paragraphs use it", () => {
    const paragraphs = Array.from({ length: 300 }, (_, i) => `<text:p text:style-name="s${19_999 - (i % 50)}">x</text:p>`).join("");
    const started = Date.now();
    limitExceeded(() => readOdt(odt(withStyles(chain(20_000), paragraphs))));
    expect(Date.now() - started).toBeLessThan(2000);
  });

  test("12 a chain at the depth cap reads, and a loop in the parents stops", () => {
    expect(readOdt(odt(withStyles(chain(50), `<text:p text:style-name="s49">x</text:p>`))).model.blocks).toHaveLength(1);
    const loop = `<style:style style:name="a" style:parent-style-name="b"/><style:style style:name="b" style:parent-style-name="a"/>`;
    expect(readOdt(odt(withStyles(loop, `<text:p text:style-name="a">x</text:p>`))).model.blocks).toHaveLength(1);
  });

  test("12 a style named like a key of Object.prototype reads as plain text", () => {
    const names = ["constructor", "toString", "hasOwnProperty", "__proto__"];
    const styles = names.map((n) => `<style:style style:name="${n}"/>`).join("");
    const body = names.map((n) => `<text:p text:style-name="${n}">x<text:span text:style-name="${n}">y</text:span></text:p>`).join("");
    const read = readOdt(odt(withStyles(styles, body)));
    expect(read.model.blocks).toHaveLength(4);
    expect(readOdt(odt(wrap(`<text:p text:style-name="constructor">x</text:p>`))).model.blocks).toHaveLength(1);
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
