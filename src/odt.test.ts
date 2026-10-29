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
 * 16 a run field of the wrong type turns a cap off; a deep or cyclic model ends in a stack overflow;
 *    one huge run is built before it is charged
 */
import { isDeepStrictEqual } from "node:util";
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, expect, test } from "vitest";
import { FormatError } from "./limits.js";
import { normalizeTextDoc, type Block, type TextDoc } from "./model.js";
import { readOdt, writeOdt } from "./odt.js";
import { PNG, textFixture as fixture } from "./testing/fixtures.js";
import { LibreOffice, hasLibreOffice, libreOfficeRequired, sha256, writeArtifact, type RoundTripRecord } from "./testing/libreoffice.js";
import { writeZip } from "./zip.js";
import { withLimits } from "./limits.js";

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

const text = (t: string) => ({ type: "text" as const, text: t });

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
    const emptyRow = Array.from({ length: 5000 }, () => [] as Block[]);
    const empty = { type: "table" as const, rows: Array.from({ length: 5000 }, () => emptyRow) };
    limitExceeded(() => writeOdt({ blocks: [empty] }, withLimits({ maxTotalBytes: 5_000_000 })));
    // Forty lists whose items each hold the list below twice: 2^40 items.
    const level = (k: number): Block => {
      const inner = k === 0 ? undefined : (level(k - 1) as Extract<Block, { type: "list" }>);
      return { type: "list", ordered: false, items: [{ content: [], ...(inner ? { sublist: inner } : {}) }, { content: [], ...(inner ? { sublist: inner } : {}) }] };
    };
    limitExceeded(() => writeOdt({ blocks: [level(40)] }, withLimits({ maxTotalBytes: 1_000_000 })));
    // Forty objects that stand for 2^40 tables: each level holds the level below twice.
    const nest = (k: number): Block => {
      if (k === 0) return { type: "paragraph", content: [{ type: "text", text: "x" }] };
      const inner = nest(k - 1);
      return { type: "table", rows: [[[inner], [inner]]] };
    };
    limitExceeded(() => writeOdt({ blocks: [nest(40)] }, withLimits({ maxTotalBytes: 1_000_000 })));
  });

  test("17 a long link or image description in a repeated table cell counts once for each copy", () => {
    const link = `<table:table ${TABLE_NS}><table:table-row table:number-rows-repeated="1000"><table:table-cell table:number-columns-repeated="1000"><text:p><text:a xlink:href="${"a".repeat(100_000)}">x</text:a></text:p></table:table-cell></table:table-row></table:table>`;
    limitExceeded(() => readOdt(odt(wrap(link))));
    const alt = `<table:table ${TABLE_NS}><table:table-row table:number-rows-repeated="1000"><table:table-cell table:number-columns-repeated="1000"><text:p><draw:frame><draw:image xlink:href="Pictures/a.png"/><svg:title xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">${"d".repeat(100_000)}</svg:title></draw:frame></text:p></table:table-cell></table:table-row></table:table>`;
    limitExceeded(() => readOdt(odt(wrap(alt), "application/vnd.oasis.opendocument.text", [{ name: "Pictures/a.png", data: PNG }])));
  });

  test("17 empty blocks in a repeated table cell are not free", () => {
    const cell = `<table:table-cell table:number-columns-repeated="1000">${"<text:p/>".repeat(1000)}</table:table-cell>`;
    const started = Date.now();
    limitExceeded(() => readOdt(odt(wrap(`<table:table ${TABLE_NS}><table:table-row table:number-rows-repeated="1000">${cell}</table:table-row></table:table>`))));
    expect(Date.now() - started).toBeLessThan(3000);
    const few = `<table:table ${TABLE_NS}><table:table-row table:number-rows-repeated="3"><table:table-cell table:number-columns-repeated="3"><text:p>x</text:p><text:p/></table:table-cell></table:table-row></table:table>`;
    expect(readOdt(odt(wrap(few))).model.blocks).toHaveLength(1);
  });

  test("17 a length attribute of a million digits is not matched, and is read as no length", () => {
    const frame = (width: string) => wrap(`<text:p><draw:frame svg:width="${width}" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><draw:image xlink:href="Pictures/a.png"/></draw:frame></text:p>`);
    const started = Date.now();
    expect(readOdt(odt(frame("9".repeat(900_000) + "x"), "application/vnd.oasis.opendocument.text", [{ name: "Pictures/a.png", data: PNG }])).model.blocks).toHaveLength(1);
    expect(Date.now() - started).toBeLessThan(2000);
  });

  test("15 a section with many blocks, and a table with many rows, fail or pass as a FormatError, never a stack overflow", { timeout: 60_000 }, () => {
    const paragraphs = Array.from({ length: 200_000 }, (_, i) => `<text:p>${i}</text:p>`).join("");
    expect(readOdt(odt(wrap(`<text:section>${paragraphs}</text:section>`))).model.blocks).toHaveLength(200_000);
    const row = [[{ type: "paragraph" as const, content: [{ type: "text" as const, text: "x" }] }]];
    const table = { type: "table" as const, rows: Array.from({ length: 200_000 }, () => row) };
    const written = writeOdt({ blocks: [table] }, withLimits({ maxTotalBytes: 100_000_000 }));
    expect(written.bytes.length).toBeGreaterThan(0);
  });

  test("16 a run whose text or link is not a string is refused, and cannot turn a cap off", () => {
    const doc = (run: unknown) => ({ blocks: [{ type: "paragraph", content: [run] }] }) as unknown as TextDoc;
    const runs = [
      { type: "text", text: Array(100).fill("a".repeat(1000)) },
      { type: "text", text: 7 },
      { type: "text", text: "x", href: { length: -1e15 } },
      { type: "text", text: "x", href: 5 },
    ];
    for (const run of runs) expect(() => writeOdt(doc(run)), JSON.stringify(run)).toThrow(/needs text and a link that are strings/);
    expect(() => writeOdt(doc(null))).toThrow(/is not an object/);
    expect(() => writeOdt(doc({ type: "bar" }))).toThrow(/is not a kind of inline content/);
  });

  test("16 a deep or cyclic model is refused with a FormatError before the stack ends", () => {
    let deepList: Block = { type: "list", ordered: false, items: [{ content: [] }] };
    for (let i = 0; i < 20_000; i++) deepList = { type: "list", ordered: false, items: [{ content: [], sublist: deepList as Extract<Block, { type: "list" }> }] };
    limitExceeded(() => writeOdt({ blocks: [deepList] }));
    const cyclicList = { type: "list" as const, ordered: false, items: [{ content: [] } as { content: never[]; sublist?: unknown }] };
    cyclicList.items[0].sublist = cyclicList;
    limitExceeded(() => writeOdt({ blocks: [cyclicList as unknown as Block] }));
    const cyclicTable = { type: "table" as const, rows: [[[] as Block[]]] };
    cyclicTable.rows[0][0].push(cyclicTable);
    limitExceeded(() => writeOdt({ blocks: [cyclicTable] }));
    expect(writeOdt({ blocks: [{ type: "list", ordered: false, items: [{ content: [], sublist: { type: "list", ordered: false, items: [{ content: [text("x")] }] } }] }] }).bytes.length).toBeGreaterThan(0);
  });

  test("16 a footnote that holds a footnote, an image, or itself is refused, not followed", () => {
    const cyclic = { type: "footnote" as const, content: [] as unknown[] };
    cyclic.content.push(cyclic);
    const nested = { type: "footnote" as const, content: [{ type: "footnote", content: [text("x")] }] };
    const withImage = { type: "footnote" as const, content: [{ type: "image", data: new Uint8Array(30), mime: "image/png" }] };
    for (const note of [cyclic, nested, withImage]) {
      expect(() => writeOdt({ blocks: [{ type: "paragraph", content: [note as never] }] })).toThrow(/a footnote holds only text and line breaks/);
    }
    let deep: unknown = { type: "footnote", content: [text("x")] };
    for (let i = 0; i < 20_000; i++) deep = { type: "footnote", content: [deep] };
    expect(() => writeOdt({ blocks: [{ type: "paragraph", content: [deep as never] }] })).toThrow(FormatError);
  });

  test("16 a model of the wrong shape is an INVALID_MODEL error, not a TypeError", () => {
    const wrong = (blocks: unknown) => ({ blocks }) as unknown as TextDoc;
    const para = (content: unknown) => wrong([{ type: "paragraph", content }]);
    for (const doc of [
      wrong([null]),
      wrong([{ type: "paragraph" }]),
      para("x"),
      para([{ type: "text", text: "x", marks: 5 }]),
      para([{ type: "image", data: "abc", mime: "image/png" }]),
      para([{ type: "image", data: null, mime: "image/png" }]),
      para([{ type: "image", data: new Uint8Array(30), mime: "image/png", alt: 5 }]),
      wrong([{ type: "list", ordered: false }]),
      wrong([{ type: "list", ordered: false, items: [null] }]),
      wrong([{ type: "table", rows: [[["x"]]] }]),
      wrong([{ type: "table", rows: [5] }]),
      wrong([{ type: "table" }]),
    ]) {
      expect(() => writeOdt(doc), JSON.stringify(doc)).toThrow(FormatError);
    }
  });

  test("16 the writer's depth cap keeps its output inside the reader's depth cap", () => {
    let deep: Block = { type: "paragraph", content: [text("x")] };
    for (let i = 0; i < 33; i++) deep = { type: "list", ordered: false, items: [{ content: [], sublist: { type: "list", ordered: false, items: [{ content: [text("x")] }] } }] };
    let table: Block = { type: "paragraph", content: [text("x")] };
    for (let i = 0; i < 32; i++) table = { type: "table", rows: [[[table]]] };
    for (const doc of [{ blocks: [deep] }, { blocks: [table] }]) expect(readOdt(writeOdt(doc).bytes).model.blocks.length).toBeGreaterThan(0);
    let tooDeep: Block = { type: "paragraph", content: [text("x")] };
    for (let i = 0; i < 40; i++) tooDeep = { type: "table", rows: [[[tooDeep]]] };
    limitExceeded(() => writeOdt({ blocks: [tooDeep] }));
  });

  test("16 an image description counts against the text cap, and a long one is refused before it is escaped", () => {
    const image = { type: "image" as const, data: PNG, mime: "image/png" as const, alt: '"'.repeat(30_000) };
    limitExceeded(() => writeOdt({ blocks: [{ type: "paragraph", content: Array.from({ length: 10 }, () => image) }] }, withLimits({ maxTotalBytes: 1_000_000 })));
    limitExceeded(() => writeOdt({ blocks: [{ type: "paragraph", content: Array.from({ length: 10 }, () => image) }] }, withLimits({ maxTextChars: 200_000 })));
  });

  test("16 one run of eight million newlines is refused before its XML is made", () => {
    const started = Date.now();
    limitExceeded(() => writeOdt({ blocks: [{ type: "paragraph", content: [{ type: "text", text: "\n".repeat(8_000_000) }] }] }));
    expect(Date.now() - started).toBeLessThan(3000);
  });

  test("15 a run of newlines is charged as the XML it becomes, before the next piece is made", () => {
    const run = { type: "text" as const, text: "\n".repeat(2_000_000) };
    limitExceeded(() => writeOdt({ blocks: [{ type: "paragraph", content: [run, run, run, run, run, run] }] }, withLimits({ maxTextChars: 20_000_000 })));
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
    const equal = isDeepStrictEqual(back.model, normalizeTextDoc(fixture));
    records.push({ format: "odt", case: "fixture", written: sha256(bytes), resaved: sha256(resaved), pdfBytes: pdf.length, equal });
    expect(back.model).toEqual(normalizeTextDoc(fixture));
  });
});
