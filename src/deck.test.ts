/*
 * ODP and PPTX, end to end: the real writers, the real readers, an independent XML parser, and
 * LibreOffice. The ways it can fail are written first; each has a case below.
 *
 *  1 text is lost or changed: marks, a link, runs of spaces, a tab, a line break, an empty
 *    paragraph, markup characters
 *  2 a shape is lost or moved: rectangle, ellipse, line, colors, no color, a picture, shape order,
 *    an empty slide
 *  3 notes are lost, or lose a line
 *  4 the deck size is lost
 *  5 the package is malformed: a part does not parse, a part is missing from the content types, a
 *    relationship points nowhere, `mimetype` is not first and stored
 *  6 the same model gives different bytes
 *  7 a writer takes a model it cannot write: no slide, sizes or colors that are not usable; an
 *    unsafe link or a bad image drops without a word; a size limit is missing
 *  8 a hostile file is read: a DOCTYPE, the wrong type, a remote or external image, a script link,
 *    a slide that is missing; a placeholder with no position takes it from its layout
 *  9 LibreOffice cannot open our file, or its re-save does not read back to the same model, in the
 *    same format or after a trip through the other format
 */
import { isDeepStrictEqual } from "node:util";
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, expect, test } from "vitest";
import { FormatError, withLimits } from "./limits.js";
import type { Deck } from "./model.js";
import { readOdp, writeOdp } from "./odp.js";
import { writeOdt } from "./odt.js";
import { readPptx, writePptx } from "./pptx.js";
import { PNG, deckFixture as fixture } from "./testing/fixtures.js";
import { LibreOffice, hasLibreOffice, libreOfficeRequired, sha256, writeArtifact, type RoundTripRecord } from "./testing/libreoffice.js";
import { writeZip } from "./zip.js";

const formats = ["odp", "pptx"] as const;
const writers = { odp: writeOdp, pptx: writePptx } as const;
const readers = { odp: readOdp, pptx: readPptx } as const;
const files = (bytes: Uint8Array) => unzipSync(bytes);
const patch = (bytes: Uint8Array, name: string, change: (xml: string) => string): Uint8Array =>
  writeZip(Object.entries(files(bytes)).map(([n, data]) => ({ name: n, data: n === name ? change(strFromU8(data)) : data, store: n === "mimetype" })), withLimits());
const wellFormed = (xml: string): void => {
  const seen: string[] = [];
  new DOMParser({ onError: (level, message) => seen.push(`${level}: ${message}`) }).parseFromString(xml, "text/xml");
  expect(seen).toEqual([]);
};
const one = (shape: Deck["slides"][0]["shapes"][0]): Deck => ({ width: 720, height: 405, slides: [{ shapes: [shape] }] });

describe.each(formats)("%s: write then read", (format) => {
  const write = writers[format];
  const read = readers[format];

  test("1-4 the deck reads back unchanged, with no warnings", () => {
    const written = write(fixture);
    const back = read(written.bytes);
    expect(back.model).toEqual(fixture);
    expect(back.warnings).toEqual([]);
    expect(written.warnings).toEqual([]);
  });

  test("2 shapes at fractional positions and with upper-case colors keep their place and color", () => {
    const deck = one({ type: "shape", kind: "rectangle", frame: { x: 10.5, y: 20.25, width: 100.7, height: 50.1 }, fill: "#FFAA00", stroke: "#00FF00" });
    const shape = read(write(deck).bytes).model.slides[0].shapes[0];
    expect(shape).toMatchObject({ type: "shape", kind: "rectangle", fill: "#ffaa00", stroke: "#00ff00", frame: { x: 10.5, y: 20.3, width: 100.7, height: 50.1 } });
  });

  test("6 the same model gives the same bytes", () => {
    expect(write(fixture).bytes).toEqual(write(fixture).bytes);
  });

  test("7 a writer refuses a model it cannot write", () => {
    const refuse = (deck: Deck) => expect(() => write(deck)).toThrow(FormatError);
    refuse({ width: 720, height: 405, slides: [] });
    refuse({ ...fixture, width: 0 });
    refuse({ ...fixture, height: Number.NaN });
    for (const n of [Number.NaN, Infinity, -1, 1e9]) refuse(one({ type: "textbox", frame: { x: 0, y: 0, width: n, height: 10 }, paragraphs: [] }));
    for (const color of ["red", "#fff", "#gggggg", "ff0000", "url(x)"]) refuse(one({ type: "shape", kind: "rectangle", frame: { x: 0, y: 0, width: 1, height: 1 }, fill: color }));
  });

  test("7 an unsafe link and a bad image are reported", () => {
    const deck: Deck = {
      width: 720,
      height: 405,
      slides: [
        {
          shapes: [
            { type: "textbox", frame: { x: 0, y: 0, width: 100, height: 20 }, paragraphs: [[{ type: "text", text: "click", href: "javascript:alert(1)" }]] },
            { type: "picture", frame: { x: 0, y: 30, width: 10, height: 10 }, data: new TextEncoder().encode("<svg onload=alert(1)/>"), mime: "image/png" },
          ],
        },
      ],
    };
    const { bytes, warnings } = write(deck);
    expect(warnings).toEqual(["a link with an unsafe target was written as plain text", "an image is not a PNG, JPEG, or GIF: not written"]);
    const all = Object.values(files(bytes)).map((d) => strFromU8(d)).join("\n");
    expect(all).toContain("click");
    expect(all).not.toMatch(/javascript|<svg/);
  });

  test("7 an image over the entry limit, and an export over the total limit, are refused", () => {
    const big = new Uint8Array(PNG.length + 4000);
    big.set(PNG);
    const deck = one({ type: "picture", frame: { x: 0, y: 0, width: 10, height: 10 }, data: big, mime: "image/png" });
    expect(() => write(deck, { maxEntryBytes: 1500 })).toThrow(FormatError);
    expect(() => write(deck, { maxTotalBytes: 6000 })).toThrow(/export is larger/);
  });

  test("10 a model that holds one string many times is refused past the text cap", () => {
    const run = { type: "text" as const, text: "n".repeat(1000) };
    const deck = one({ type: "textbox", frame: { x: 0, y: 0, width: 10, height: 10 }, paragraphs: Array.from({ length: 50 }, () => [run]) });
    expect(() => write(deck, { maxTextChars: 10_000 })).toThrow(/more than 10000 characters of text/);
    const notes = "n".repeat(1000);
    expect(() => write({ width: 720, height: 405, slides: Array.from({ length: 50 }, () => ({ shapes: [], notes })) }, { maxTextChars: 10_000 })).toThrow(/characters of text/);
    expect(() => write(deck)).not.toThrow();
  });

  test("5 every XML part parses with an independent parser", () => {
    const all = files(write(fixture).bytes);
    const xml = Object.keys(all).filter((n) => /\.(xml|rels)$/.test(n));
    expect(xml.length).toBeGreaterThanOrEqual(4);
    for (const name of xml) wellFormed(strFromU8(all[name]));
  });
});

describe("odp package", () => {
  test("5 `mimetype` is first and stored, and the manifest carries `manifest:version`", () => {
    const { bytes } = writeOdp(fixture);
    expect(strFromU8(bytes.subarray(30, 38))).toBe("mimetype");
    expect(new DataView(bytes.buffer, bytes.byteOffset).getUint16(8, true)).toBe(0);
    expect(strFromU8(files(bytes)["META-INF/manifest.xml"])).toMatch(/<manifest:manifest [^>]*manifest:version="1\.3"/);
  });
});

describe("pptx package", () => {
  test("5 every part is in the content types, and every internal relationship target exists", () => {
    const all = files(writePptx(fixture).bytes);
    const types = strFromU8(all["[Content_Types].xml"]);
    for (const part of Object.keys(all).filter((n) => n.endsWith(".xml") && n !== "[Content_Types].xml")) expect(types).toContain(`PartName="/${part}"`);
    expect(Object.keys(all)[0]).toBe("[Content_Types].xml");
    for (const rel of Object.keys(all).filter((n) => n.endsWith(".rels"))) {
      const dir = rel.replace(/_rels\/[^/]+$/, "");
      for (const m of strFromU8(all[rel]).matchAll(/Target="([^"]+)"(?! TargetMode)/g)) {
        const path = [...dir.split("/").filter(Boolean), ...m[1].split("/")].reduce<string[]>((acc, p) => (p === ".." ? (acc.pop(), acc) : [...acc, p]), []).join("/");
        expect(all[path], `${rel} -> ${m[1]}`).toBeDefined();
      }
    }
  });

  test("8 a placeholder with no position takes it from its layout", () => {
    const base = writePptx({ width: 720, height: 405, slides: [{ shapes: [{ type: "textbox", frame: { x: 0, y: 0, width: 1, height: 1 }, paragraphs: [[{ type: "text", text: "Title" }]] }] }] }).bytes;
    const layout = patch(base, "ppt/slideLayouts/slideLayout1.xml", (x) =>
      x.replace("</p:spTree>", `<p:sp><p:nvSpPr><p:cNvPr id="9" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${36 * 12700}" y="${18 * 12700}"/><a:ext cx="${600 * 12700}" cy="${50 * 12700}"/></a:xfrm></p:spPr></p:sp></p:spTree>`),
    );
    const slide = patch(layout, "ppt/slides/slide1.xml", (x) => x.replace('<p:cNvSpPr txBox="1"/><p:nvPr/>', '<p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr>').replace(/<p:spPr>.*?<\/p:spPr>/, "<p:spPr/>"));
    const read = readPptx(slide);
    expect(read.model.slides[0].shapes[0]).toMatchObject({ type: "textbox", frame: { x: 36, y: 18, width: 600, height: 50 } });
    expect(read.warnings).toEqual([]);
  });
});

describe("slow files", () => {
  const frame = { x: 0, y: 0, width: 100, height: 50 };
  const slideOf = (text: string, notes?: string) => ({ shapes: [{ type: "textbox" as const, frame, paragraphs: [[{ type: "text" as const, text }]] }], ...(notes ? { notes } : {}) });

  test("10 500 slides that share one big layout read fast, and each placeholder takes its place from it", { timeout: 120_000 }, () => {
    const deck: Deck = { width: 720, height: 405, slides: Array.from({ length: 500 }, (_, i) => slideOf(`slide ${i}`)) };
    const placeholders = Array.from(
      { length: 10_000 },
      (_, i) => `<p:sp><p:nvSpPr><p:cNvPr id="${i + 2}" name="P${i}"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="${i}"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.floor(i / 2) * 12700}" y="0"/><a:ext cx="12700" cy="12700"/></a:xfrm></p:spPr></p:sp>`,
    ).join("");
    // The master holds 10,000 placeholders of its own (idx 10000 and up); odd slides take theirs from it.
    const masterPlaceholders = placeholders
      .replaceAll(/idx="(\d+)"/g, (_m, i: string) => `idx="${Number(i) + 10_000}"`)
      .replaceAll(/<a:off x="(\d+)"/g, (_m, x: string) => `<a:off x="${Number(x) + 8000 * 12_700}"`);
    const entries = Object.entries(files(writePptx(deck).bytes)).map(([name, data]) => {
      if (name === "ppt/slideLayouts/slideLayout1.xml") return { name, data: strFromU8(data).replace("</p:spTree>", `${placeholders}</p:spTree>`) };
      if (name === "ppt/slideMasters/slideMaster1.xml") return { name, data: strFromU8(data).replace("</p:spTree>", `${masterPlaceholders}</p:spTree>`) };
      const slide = /^ppt\/slides\/slide(\d+)\.xml$/.exec(name);
      if (slide) {
        const idx = Number(slide[1]) % 2 ? 9999 : 19_999;
        return { name, data: strFromU8(data).replace('<p:cNvSpPr txBox="1"/><p:nvPr/>', `<p:cNvSpPr/><p:nvPr><p:ph type="body" idx="${idx}"/></p:nvPr>`).replace(/(<p:spPr>)<a:xfrm>.*?<\/a:xfrm>/, "$1") };
      }
      return { name, data };
    });
    const started = Date.now();
    const read = readPptx(writeZip(entries, withLimits()));
    // Parsing each shared part once takes a few seconds; parsing it for every slide took 174 seconds.
    expect(Date.now() - started).toBeLessThan(20_000);
    expect(read.model.slides).toHaveLength(500);
    expect(read.model.slides[0].shapes[0]).toMatchObject({ frame: { x: 4999, y: 0 } });
    expect(read.model.slides[1].shapes[0]).toMatchObject({ frame: { x: 12_999, y: 0 } });
    expect(read.warnings).toEqual([]);
  });

  test("10 a slide part or a notes part that two entries share is read once, and the repeat warns", () => {
    const base = writePptx({ width: 720, height: 405, slides: [slideOf("one", "notes one"), slideOf("two", "notes two")] }).bytes;
    const twice = patch(base, "ppt/presentation.xml", (x) => x.replace('<p:sldId id="257" r:id="rId3"/>', '<p:sldId id="257" r:id="rId2"/>'));
    const slides = readPptx(twice);
    expect(slides.model.slides).toHaveLength(1);
    expect(slides.warnings).toEqual(["a slide that is listed twice: not imported"]);
    const shared = patch(base, "ppt/slides/_rels/slide2.xml.rels", (x) => x.replace("notesSlide2.xml", "notesSlide1.xml"));
    const notes = readPptx(shared);
    expect(notes.model.slides.map((s) => s.notes)).toEqual(["notes one", undefined]);
    expect(notes.warnings).toEqual(["a notes page that two slides share: not imported"]);
  });

  test("11 one run or one note of eight million newlines is refused before its XML is made, in both writers", { timeout: 60_000 }, () => {
    const deck = (shape?: Deck["slides"][0]["shapes"][0], notes?: string): Deck => ({ width: 720, height: 405, slides: [{ shapes: shape ? [shape] : [], ...(notes ? { notes } : {}) }] });
    const run = { type: "text" as const, text: "\n".repeat(8_000_000) };
    for (const write of [writePptx, writeOdp]) {
      for (const model of [deck({ type: "textbox", frame, paragraphs: [[run]] }), deck(undefined, "\n".repeat(8_000_000))]) {
        const started = Date.now();
        expect(() => write(model)).toThrow(/passes 104857600 bytes while it is written/);
        expect(Date.now() - started).toBeLessThan(3000);
      }
    }
  });

  test("11 millions of empty runs that a model holds are charged for the string they make", { timeout: 60_000 }, () => {
    const empty = { type: "text" as const, text: "", href: "a" };
    const model: Deck = { width: 720, height: 405, slides: [{ shapes: [{ type: "textbox", frame, paragraphs: [Array.from({ length: 4_000_000 }, () => empty)] }] }] };
    const started = Date.now();
    expect(() => writePptx(model)).toThrow(/while it is written/);
    expect(Date.now() - started).toBeLessThan(5000);
  });

  test("11 an empty linked run writes no relationship, and an in-page link is reported", () => {
    const linked = writePptx({ width: 720, height: 405, slides: [{ shapes: [{ type: "textbox", frame, paragraphs: [[{ type: "text", text: "x", href: "#top" }, { type: "text", text: "", href: "https://example.org/e" }]] }] }] });
    expect(linked.warnings).toEqual(["an in-page link was written as plain text"]);
    expect(strFromU8(files(linked.bytes)["ppt/slides/_rels/slide1.xml.rels"])).not.toContain("example.org");
  });

  test("11 a model of the wrong shape is an INVALID_MODEL error, not a TypeError", () => {
    const wrong = (slide: unknown) => ({ width: 720, height: 405, slides: [slide] }) as unknown as Deck;
    const shape = (s: unknown) => wrong({ shapes: [s] });
    for (const model of [{ width: 720, height: 405, slides: {} } as unknown as Deck, wrong(null), wrong({ shapes: [{ type: "text", frame, paragraphs: [] }] }), wrong({ shapes: [{ type: "shape", kind: "rectangle" }] }), shape({ type: "shape", kind: "rectangle", frame, fill: ["#ff0000"] }), shape({ type: "shape", kind: "rectangle", frame, stroke: 5 }), shape({ type: "shape", kind: "circle", frame }), shape({ type: "shape", kind: 7, frame }), shape({ type: "shape", frame }), wrong({}), wrong({ shapes: [], notes: 7 }), shape(null), shape({ type: "textbox", frame, paragraphs: "x" }), shape({ type: "textbox", frame, paragraphs: [[{ type: "text", text: 7 }]] }), shape({ type: "picture", frame, data: PNG, mime: "image/png", alt: 5 })]) {
      for (const write of [writePptx, writeOdp]) expect(() => write(model)).toThrow(FormatError);
    }
  });

  test("11 each run that names a shared link target counts it, in the PPTX and ODP readers", { timeout: 60_000 }, () => {
    const target = `https://example.org/${"u".repeat(100_000)}`;
    const deck = writePptx({ width: 720, height: 405, slides: [{ shapes: [{ type: "textbox", frame, paragraphs: [[{ type: "text", text: "a", href: "https://example.org/x" }]] }] }] }).bytes;
    const run = `<a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId3"/></a:rPr><a:t>a</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t>b</a:t></a:r>`;
    const pptx = (runs: number) =>
      patch(patch(deck, "ppt/slides/_rels/slide1.xml.rels", (x) => x.replace(/Target="https:\/\/example\.org\/x"/, `Target="${target}"`)), "ppt/slides/slide1.xml", (x) => x.replace(/(<a:p>).*(<\/a:p>)/, `$1${run.repeat(runs)}$2`));
    expect(readPptx(pptx(25)).model.slides).toHaveLength(1);
    expect(() => readPptx(pptx(200))).toThrow(/characters of run text and link targets/);
    const odp = writeOdp({ width: 720, height: 405, slides: [{ shapes: [{ type: "textbox", frame, paragraphs: [[{ type: "text", text: "a" }]] }] }] }).bytes;
    // One link around many spans of two styles: each leaf carries the same target, and unlike marks stop the runs from merging.
    const bold = `<style:style style:name="B" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>`;
    const runs = (n: number) =>
      patch(odp, "content.xml", (x) =>
        x.replace("</office:automatic-styles>", `${bold}</office:automatic-styles>`).replace(/<text:p text:style-name="Pdeck">.*?<\/text:p>/, `<text:p text:style-name="Pdeck"><text:a xlink:href="${target}">${'<text:span text:style-name="B">a</text:span>b'.repeat(n)}</text:a></text:p>`),
      );
    expect(readOdp(runs(25)).model.slides).toHaveLength(1);
    expect(() => readOdp(runs(200))).toThrow(/characters of run text and link targets/);
  });

  test("11 an ODP line outside the bounds, and a deck of too many slides, are not read", () => {
    const odp = writeOdp({ width: 720, height: 405, slides: [slideOf("x")] }).bytes;
    const far = readOdp(patch(odp, "content.xml", (x) => x.replace("</draw:page>", `<draw:line svg:x1="-${"9".repeat(30)}cm" svg:y1="0cm" svg:x2="${"9".repeat(30)}cm" svg:y2="0cm"/></draw:page>`)));
    expect(far.model.slides[0].shapes[1]).toMatchObject({ kind: "line", frame: { x: 0, y: 0, width: 0, height: 0 } });
    expect(far.warnings).toEqual(["a shape with a position or size outside the slide bounds was placed at the top left corner"]);
    const page = (n: number) => `<draw:page draw:name="p${n}" draw:master-page-name="Default"/>`;
    const many = patch(odp, "content.xml", (x) => x.replace("</office:presentation>", `${Array.from({ length: 1300 }, (_, i) => page(i)).join("")}</office:presentation>`));
    expect(() => readOdp(many)).toThrow(/more than 1250 slides/);
    const pptx = writePptx({ width: 720, height: 405, slides: [slideOf("x")] }).bytes;
    const listed = patch(pptx, "ppt/presentation.xml", (x) => x.replace("</p:sldIdLst>", `${'<p:sldId id="300" r:id="rId2"/>'.repeat(1300)}</p:sldIdLst>`));
    expect(() => readPptx(listed)).toThrow(/more than 1250 slides/);
  });

  test("11 an ODP shape or deck size outside the bounds is not read", () => {
    const odp = writeOdp({ width: 720, height: 405, slides: [slideOf("x")] }).bytes;
    const wide = readOdp(patch(odp, "content.xml", (x) => x.replace('svg:width="3.528cm"', 'svg:width="-5cm"').replace(/svg:width="[^"]*"/, 'svg:width="-5cm"')));
    expect(wide.model.slides[0].shapes[0]).toMatchObject({ frame: { x: 0, y: 0, width: 0, height: 0 } });
    expect(wide.warnings).toEqual(["a shape with a position or size outside the slide bounds was placed at the top left corner"]);
    const size = readOdp(patch(odp, "styles.xml", (x) => x.replace(/fo:page-width="[^"]*"/, 'fo:page-width="-10cm"')));
    expect(size.model.width).toBe(720);
    expect(size.warnings).toEqual(["the slide size was missing: 720 by 405 points assumed"]);
  });

  test("11 a size or position that is not a number, or is out of bounds, is not read from a PPTX", () => {
    const base = writePptx({ width: 720, height: 405, slides: [slideOf("x")] }).bytes;
    const box = (attr: "off x" | "ext cx", value: string) => (x: string) => x.replace(attr === "off x" ? /(<p:spPr><a:xfrm><a:off x=")\d+"/ : /(<p:spPr><a:xfrm><a:off [^>]*\/><a:ext cx=")\d+"/, `$1${value}"`);
    for (const [attr, value] of [["off x", "abc"], ["ext cx", "-127000"], ["ext cx", "1e30"]] as const) {
      const read = readPptx(patch(base, "ppt/slides/slide1.xml", box(attr, value)));
      expect(read.warnings, `${attr} ${value}`).toEqual(["a shape without a position was placed at the top left corner"]);
    }
    expect(() => readPptx(patch(base, "ppt/presentation.xml", (x) => x.replace(/<p:sldSz cx="\d+"/, '<p:sldSz cx="999999999999999"')))).toThrow(/no slide size/);
  });

  test("11 a connector with no position is reported", () => {
    const base = writePptx({ width: 720, height: 405, slides: [slideOf("x")] }).bytes;
    const read = readPptx(patch(base, "ppt/slides/slide1.xml", (x) => x.replace("</p:spTree>", `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="9" name="c"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr/></p:cxnSp></p:spTree>`)));
    expect(read.warnings).toEqual(["a connector without a position: not imported"]);
  });

  test("11 a link target and a description that many runs or pictures share count against the text cap", () => {
    const link = { type: "text" as const, text: "x", href: `https://example.org/${"a".repeat(10_000)}` };
    const deck = (n: number): Deck => ({ width: 720, height: 405, slides: [{ shapes: [{ type: "textbox", frame, paragraphs: [Array.from({ length: n }, () => link)] }] }] });
    for (const write of [writePptx, writeOdp]) {
      expect(() => write(deck(5), { maxTextChars: 100_000 })).not.toThrow();
      expect(() => write(deck(20), { maxTextChars: 100_000 })).toThrow(/more than 100000 characters of text/);
    }
    const picture = { type: "picture" as const, frame, data: PNG, mime: "image/png" as const, alt: "d".repeat(10_000) };
    expect(() => writePptx({ width: 720, height: 405, slides: [{ shapes: Array.from({ length: 20 }, () => picture) }] }, { maxTextChars: 100_000 })).toThrow(/characters of text/);
  });

  test("11 one shape or one empty paragraph that a model holds many times stops at the byte cap", () => {
    const shape = { type: "shape" as const, kind: "rectangle" as const, frame };
    const many = { width: 720, height: 405, slides: [{ shapes: Array.from({ length: 200_000 }, () => shape) }] };
    const empty = { width: 720, height: 405, slides: [{ shapes: [{ type: "textbox" as const, frame, paragraphs: Array.from({ length: 500_000 }, () => [] as never[]) }] }] };
    for (const write of [writePptx, writeOdp]) {
      for (const deck of [many, empty]) {
        const started = Date.now();
        expect(() => write(deck, { maxTotalBytes: 1_000_000 })).toThrow(/passes 1000000 bytes while it is written/);
        expect(Date.now() - started).toBeLessThan(5000);
      }
    }
  });

  test("11 text:s runs count against the text cap, and nested spans keep seven marks at most", () => {
    const odp = writeOdp({ width: 720, height: 405, slides: [slideOf("x")] }).bytes;
    const spaces = (n: number) => patch(odp, "content.xml", (x) => x.replace(">x</text:p>", `>a${'<text:s text:c="1000"/>'.repeat(n)}b</text:p>`));
    expect(() => readOdp(spaces(11), { maxTextChars: 10_000 })).toThrow(/spaces in <text:s> elements|characters of run text/);
    expect(readOdp(spaces(9), { maxTextChars: 10_000 }).warnings).toEqual([]);
    const props = `<style:text-properties fo:font-weight="bold" fo:font-style="italic" style:text-underline-style="solid" style:text-line-through-style="solid" fo:font-family="monospace" style:text-position="super 58%"/>`;
    const styles = Array.from({ length: 50 }, (_, i) => `<style:style style:name="m${i}" style:family="text"${i ? ` style:parent-style-name="m${i - 1}"` : ""}>${props}</style:style>`).join("");
    const spans = `<text:span text:style-name="m49">`.repeat(60) + "x" + "</text:span>".repeat(60);
    const started = Date.now();
    const read = readOdp(patch(odp, "content.xml", (x) => x.replace("</office:automatic-styles>", `${styles}</office:automatic-styles>`).replace(">x</text:p>", `>${spans}</text:p>`)));
    expect(Date.now() - started).toBeLessThan(2000);
    expect(read.model.slides[0].shapes[0]).toMatchObject({ paragraphs: [[{ text: "x", marks: ["bold", "italic", "underline", "strike", "code", "superscript"] }]] });
  });

  test("11 one image part named by many pictures counts its bytes for each picture", () => {
    const padded = new Uint8Array(10_000);
    padded.set(PNG);
    const picture = { type: "picture" as const, frame, data: padded, mime: "image/png" as const };
    const one = writePptx({ width: 720, height: 405, slides: [{ shapes: [picture] }] }).bytes;
    const many = patch(one, "ppt/slides/slide1.xml", (x) => x.replace(/<p:pic>.*<\/p:pic>/, (m) => m.repeat(5)));
    expect(readPptx(many, { maxTotalBytes: 60_000 }).model.slides[0].shapes).toHaveLength(5);
    expect(() => readPptx(many, { maxTotalBytes: 40_000 })).toThrow(/images in the file add up to more than 40000 bytes/);
  });

  test("10 an ODP style chain of twenty thousand styles is refused fast, a chain at the cap reads, a loop stops", () => {
    const odp = writeOdp({ width: 720, height: 405, slides: [slideOf("x")] }).bytes;
    const chain = (length: number) => Array.from({ length }, (_, i) => `<style:style style:name="s${i}" style:family="graphic"${i ? ` style:parent-style-name="s${i - 1}"` : ""}/>`).join("");
    const rects = (name: string, n: number) => Array.from({ length: n }, () => `<draw:rect draw:style-name="${name}" svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"/>`).join("");
    const withStyles = (styles: string, shapes: string) => patch(odp, "content.xml", (x) => x.replace("</office:automatic-styles>", `${styles}</office:automatic-styles>`).replace("</draw:page>", `${shapes}</draw:page>`));
    const started = Date.now();
    expect(() => readOdp(withStyles(chain(20_000), rects("s19999", 300)))).toThrow(/more than 50 parents/);
    expect(Date.now() - started).toBeLessThan(2000);
    expect(readOdp(withStyles(chain(50), rects("s49", 1))).model.slides[0].shapes).toHaveLength(2);
    const loop = `<style:style style:name="a" style:family="graphic" style:parent-style-name="b"/><style:style style:name="b" style:family="graphic" style:parent-style-name="a"/>`;
    expect(readOdp(withStyles(loop, rects("a", 1))).model.slides[0].shapes).toHaveLength(2);
  });
});

describe("hostile files", () => {
  const odp = writeOdp({ width: 720, height: 405, slides: [{ shapes: [{ type: "textbox", frame: { x: 0, y: 0, width: 10, height: 10 }, paragraphs: [[{ type: "text", text: "x" }]] }] }] }).bytes;
  const pptx = writePptx({ width: 720, height: 405, slides: [{ shapes: [{ type: "picture", frame: { x: 0, y: 0, width: 10, height: 10 }, data: PNG, mime: "image/png" }] }] }).bytes;

  test("8 a DOCTYPE with an entity is refused", () => {
    expect(() => readOdp(patch(odp, "content.xml", (x) => x.replace("?>", `?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]>`)))).toThrow(/DOCTYPE/);
    expect(() => readPptx(patch(pptx, "ppt/slides/slide1.xml", (x) => x.replace("?>", `?><!DOCTYPE x [<!ENTITY e "x">]>`)))).toThrow(/DOCTYPE/);
  });

  test("8 the wrong type of file is refused", () => {
    expect(() => readOdp(writeOdt({ blocks: [] }).bytes)).toThrow(/not an ODP/);
    expect(() => readOdp(pptx)).toThrow(FormatError);
    expect(() => readPptx(odp)).toThrow(FormatError);
    expect(() => readPptx(patch(pptx, "[Content_Types].xml", (x) => x.replace("presentationml.presentation.main", "wordprocessingml.document.main")))).toThrow(/not an PPTX/);
    expect(() => readPptx(new TextEncoder().encode("PK nope"))).toThrow(FormatError);
  });

  test("8 an external image and a script link are not followed", () => {
    const rels = patch(pptx, "ppt/slides/_rels/slide1.xml.rels", (x) => x.replace(/Target="\.\.\/media\/image1\.png"/, 'Target="https://evil.example/p.png" TargetMode="External"'));
    const read = readPptx(rels);
    expect(read.model.slides[0].shapes).toEqual([]);
    expect(read.warnings).toEqual(["an image that is not a PNG, JPEG, or GIF: not imported"]);
    const remote = readOdp(patch(odp, "content.xml", (x) => x.replace("<draw:text-box>", '<draw:image xlink:href="https://evil.example/p.png"/><draw:text-box>')));
    expect(remote.model.slides[0].shapes).toHaveLength(1);
    const link = readOdp(patch(odp, "content.xml", (x) => x.replace(">x</text:p>", '><text:a xlink:href="javascript:alert(1)">x</text:a></text:p>')));
    expect(link.warnings).toEqual(["a link with an unsafe target was read as plain text"]);
    expect(link.model.slides[0].shapes[0]).toMatchObject({ paragraphs: [[{ text: "x" }]] });
  });

  test("8 a slide that points nowhere is skipped with a warning, and a deck with none is refused", () => {
    const gone = patch(pptx, "ppt/_rels/presentation.xml.rels", (x) => x.replace('Target="slides/slide1.xml"', 'Target="slides/none.xml"'));
    expect(() => readPptx(gone)).toThrow(/no slide/);
    expect(() => readOdp(patch(odp, "content.xml", (x) => x.replace(/<draw:page .*<\/draw:page>/, "")))).toThrow(/no slide/);
  });
});

describe.skipIf(!hasLibreOffice && !libreOfficeRequired)("LibreOffice round trip", () => {
  const lo = hasLibreOffice ? new LibreOffice() : undefined;
  const records: RoundTripRecord[] = [];
  afterAll(() => {
    lo?.dispose();
    writeArtifact("deck-roundtrip", records);
  });

  test.each(formats)("9 LibreOffice opens our %s, prints it, and its re-save reads back to the same model", { timeout: 300_000 }, (format) => {
    expect(hasLibreOffice).toBe(true);
    const { bytes } = writers[format](fixture);
    const pdf = lo!.convert(bytes, format, "pdf");
    expect(strFromU8(pdf.subarray(0, 5))).toBe("%PDF-");
    const resaved = lo!.convert(bytes, format, format);
    const back = readers[format](resaved);
    records.push({ format, case: "fixture", written: sha256(bytes), resaved: sha256(resaved), pdfBytes: pdf.length, equal: isDeepStrictEqual(back.model, fixture) });
    expect(back.model).toEqual(fixture);
  });

  test.each([
    ["pptx", "odp"],
    ["odp", "pptx"],
  ] as const)("9 %s through LibreOffice to %s reads back to the same model", { timeout: 300_000 }, (from, to) => {
    const { bytes } = writers[from](fixture);
    const converted = lo!.convert(bytes, from, to);
    const back = readers[to](converted);
    records.push({ format: `${from}->${to}`, case: "fixture", written: sha256(bytes), resaved: sha256(converted), pdfBytes: 0, equal: isDeepStrictEqual(back.model, fixture) });
    expect(back.model).toEqual(fixture);
  });
});
