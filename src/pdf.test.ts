/*
 * PDF writer, end to end: the real writer, an independent check of the file structure, and pdftotext
 * when it is installed. The ways it can fail are written first; each has a case below.
 *
 *  1 the file is not a PDF: header, cross-reference table, offsets, trailer, page count
 *  2 text is lost or garbled: words, headings, lists, table cells, footnotes, characters outside WinAnsi
 *  3 a slide is not one page, or a page has the wrong size
 *  4 an image is dropped or breaks the file: PNG, PNG with alpha, JPEG, a GIF, bytes that are no image
 *  5 text that does not fit is cut without a word
 *  6 the same model gives different bytes
 *  7 a footnote holds an image or a footnote; the text limit is charged twice or counts image bytes; many
 *    PNGs with alpha cost minutes; a PNG loses transparency; a deep list runs off the page; a seeded
 *    mutation of a model throws something other than a FormatError
 *  8 a writer takes a model it cannot write: no slides, nesting too deep, a bad block; a limit is missing
 */
import { spawnSync } from "node:child_process";
import { mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { strFromU8, unzlibSync, zlibSync } from "fflate";
import { describe, expect, test } from "vitest";
import { FormatError } from "./limits.js";
import type { Deck, TextDoc } from "./model.js";
import { writePdf } from "./pdf.js";
import { PNG, deckFixture, textFixture } from "./testing/fixtures.js";

const text = (t: string, marks?: ("bold" | "italic")[]) => ({ type: "text" as const, text: t, ...(marks ? { marks } : {}) });
const para = (t: string) => ({ type: "paragraph" as const, content: [text(t)] });
const latin = (b: Uint8Array) => Array.from(b, (c) => String.fromCharCode(c)).join("");

/** Check the structure and return the page count. */
function structure(bytes: Uint8Array): number {
  const s = latin(bytes);
  expect(s.startsWith("%PDF-1.5\n")).toBe(true);
  expect(s.endsWith("%%EOF\n")).toBe(true);
  const start = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(s)![1]);
  expect(s.slice(start, start + 5)).toBe("xref\n");
  const size = Number(/^xref\n0 (\d+)\n/.exec(s.slice(start))![1]);
  const rows = s.slice(start).split("\n").slice(2, 2 + size);
  expect(rows[0]).toBe("0000000000 65535 f ");
  rows.slice(1).forEach((row, i) => {
    expect(row).toMatch(/^\d{10} 00000 n $/);
    expect(s.slice(Number(row.slice(0, 10))).startsWith(`${i + 1} 0 obj\n`)).toBe(true);
  });
  expect(s).toContain(`/Size ${size}`);
  return Number(/\/Type \/Pages \/Kids \[[^\]]*\] \/Count (\d+)/.exec(s)![1]);
}

/** The text of every page, from the compressed content streams. */
function contentText(bytes: Uint8Array): string {
  const s = latin(bytes);
  const out: string[] = [];
  for (const m of s.matchAll(/\/Filter \/FlateDecode \/Length (\d+)>>\nstream\n/g)) {
    const at = m.index! + m[0].length;
    const body = strFromU8(unzlibSync(bytes.subarray(at, at + Number(m[1]))), true);
    for (const t of body.matchAll(/\(((?:[^\\)]|\\.)*)\) Tj/g)) out.push(t[1].replace(/\\(\d{3})/g, (_, o) => String.fromCharCode(parseInt(o, 8))).replace(/\\(.)/g, "$1"));
  }
  return out.join("");
}

const hasPdftotext = spawnSync("pdftotext", ["-v"]).error === undefined;
function pdftotext(bytes: Uint8Array): string {
  const dir = mkdtempSync(join(tmpdir(), "pdf-"));
  try {
    writeFileSync(join(dir, "a.pdf"), bytes);
    const r = spawnSync("pdftotext", ["-layout", join(dir, "a.pdf"), "-"], { encoding: "utf8" });
    expect(r.status).toBe(0);
    return r.stdout;
  } finally {
    rmSync(dir, { recursive: true, force: true });
  }
}

const slideBox = (t: string, w = 400, h = 100) => ({ type: "textbox" as const, frame: { x: 20, y: 20, width: w, height: h }, paragraphs: [[text(t)]] });
const deckOf = (...slides: string[]): Deck => ({ width: 720, height: 405, slides: slides.map((t) => ({ shapes: [slideBox(t)] })) });

describe("PDF writer", () => {
  test("1 a deck is a valid PDF with one page for each slide", () => {
    const { bytes, warnings } = writePdf(deckOf("one", "two", "three"));
    expect(structure(bytes)).toBe(3);
    expect(latin(bytes)).toMatch(/\/MediaBox \[0 0 720 405\]/);
    expect(contentText(bytes)).toBe("onetwothree");
    expect(warnings).toEqual([]);
  });

  test("1 a document is a valid PDF and runs over onto a second page", () => {
    const blocks = Array.from({ length: 120 }, (_, i) => para(`Paragraph number ${i + 1} of the long document.`));
    const { bytes } = writePdf({ title: "Long", blocks });
    expect(structure(bytes)).toBeGreaterThan(1);
    expect(latin(bytes)).toMatch(/\/Title <FEFF004c006f006e0067>/);
    expect(writePdf({ blocks: [] }).bytes.length).toBeGreaterThan(0);
    expect(structure(writePdf({ blocks: [] }).bytes)).toBe(1);
  });

  test("2 headings, marks, lists, tables, and footnotes keep their words", () => {
    const { bytes, warnings } = writePdf(textFixture);
    structure(bytes);
    const t = contentText(bytes);
    for (const word of ["Plan", "Details", "bold", "italic and underline", "two-b", "deep", "first", "Apples", "paragraphs", "See", "Claim"]) expect(t).toContain(word);
    expect(warnings.some((w) => /superscript/.test(w))).toBe(true);
    if (hasPdftotext) {
      const out = pdftotext(bytes);
      expect(out).toMatch(/Plan/);
      expect(out).toMatch(/Plain, bold italic and underline struck code\(\)/);
      expect(out).toMatch(/Apples/);
      expect(out).toMatch(/1\. first/);
      expect(out).toMatch(/\[1\]/);
    }
  });

  test("2 pdftotext reads each slide's text in order", () => {
    if (!hasPdftotext) return;
    const out = pdftotext(writePdf(deckOf("Hello slide", "Second slide")).bytes);
    expect(out.indexOf("Hello slide")).toBeGreaterThanOrEqual(0);
    expect(out.indexOf("Hello slide")).toBeLessThan(out.indexOf("Second slide"));
  });

  test("2 characters outside WinAnsi are written as ? and reported once", () => {
    const { bytes, warnings } = writePdf({ blocks: [para("café €5 “quoted” ✓ ✓ 中")] });
    const t = contentText(bytes);
    expect(t).toContain("café \u0080");
    expect(t).toContain("? ?");
    expect(warnings.filter((w) => /cannot hold/.test(w))).toHaveLength(1);
    structure(bytes);
  });

  test("2 parentheses, backslashes, and a very long word are escaped and split", () => {
    const { bytes } = writePdf({ blocks: [para(`a (b) c\\d ${"x".repeat(300)}`)] });
    structure(bytes);
    expect(contentText(bytes)).toContain("a (b) c\\d ");
    expect(contentText(bytes).replace(/[^x]/g, "").length).toBe(300);
  });

  test("3 shapes and a link write drawing operators and a link annotation", () => {
    const deck = deckFixture;
    const { bytes } = writePdf(deck);
    structure(bytes);
    const doc = writePdf({ blocks: [{ type: "paragraph", content: [{ type: "text", text: "go", href: "https://example.org/a" }, { type: "text", text: "rel", href: "page.html" }] }] });
    expect(latin(doc.bytes)).toContain("/URI (https://example.org/a)");
    expect(latin(doc.bytes)).not.toContain("page.html");
    expect(doc.warnings.some((w) => /not a web or mail address/.test(w))).toBe(true);
  });

  test("4 PNG, PNG with alpha, and JPEG are embedded; a GIF and junk are reported", () => {
    const rgba = (() => {
      const crc = (b: Uint8Array) => {
        let c = ~0;
        for (const x of b) {
          c ^= x;
          for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
        }
        return ~c >>> 0;
      };
      const chunk = (type: string, body: number[]) => {
        const b = new Uint8Array(12 + body.length);
        const v = new DataView(b.buffer);
        v.setUint32(0, body.length);
        b.set([...type].map((c) => c.charCodeAt(0)), 4);
        b.set(body, 8);
        v.setUint32(8 + body.length, crc(b.subarray(4, 8 + body.length)));
        return b;
      };
      const ihdr = [0, 0, 0, 2, 0, 0, 0, 2, 8, 6, 0, 0, 0];
      // Row one has no filter; row two uses the Sub filter, so each byte is a difference from the pixel on its left.
      const idat = Array.from(zlibSync(Uint8Array.of(0, 255, 0, 0, 255, 0, 255, 0, 128, 1, 0, 0, 255, 255, 0, 0, 0, 0)));
      return new Uint8Array([...Uint8Array.of(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a), ...chunk("IHDR", ihdr), ...chunk("IDAT", idat), ...chunk("IEND", [])]);
    })();
    const jpeg = Uint8Array.of(0xff, 0xd8, 0xff, 0xc0, 0, 17, 8, 0, 4, 0, 4, 3, 1, 0x22, 0, 2, 0x11, 1, 3, 0x11, 1, 0xff, 0xd9);
    const gif = Uint8Array.of(0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 1, 0, 1, 0, 0, 0, 0, 0x3b);
    const img = (data: Uint8Array, mime: "image/png" | "image/jpeg" | "image/gif") => ({ type: "image" as const, data, mime, width: 40, height: 40 });
    const { bytes, warnings } = writePdf({ blocks: [{ type: "paragraph", content: [img(PNG, "image/png"), img(rgba, "image/png"), img(jpeg, "image/jpeg"), img(gif, "image/gif"), img(Uint8Array.of(1, 2, 3, 4), "image/png")] }] });
    structure(bytes);
    const s = latin(bytes);
    expect(s.match(/\/Subtype \/Image/g)).toHaveLength(5); // the PNG with alpha and the 1x1 fixture each add a mask
    expect(s).toContain("/SMask");
    expect(s).toContain("/DCTDecode");
    expect(warnings.some((w) => /GIF/.test(w))).toBe(true);
    expect(warnings.some((w) => /not a PNG, JPEG, or GIF/.test(w))).toBe(true);
  });

  test("4 a deck picture is drawn in its frame, and a repeated image is stored once", () => {
    const pic = { type: "picture" as const, data: PNG, mime: "image/png" as const, frame: { x: 10, y: 10, width: 100, height: 50 } };
    const { bytes } = writePdf({ width: 300, height: 200, slides: [{ shapes: [pic, pic] }, { shapes: [pic] }] });
    structure(bytes);
    // One RGBA picture is an image and its mask, however often it is drawn.
    expect(latin(bytes).match(/\/Subtype \/Image/g)).toHaveLength(2);
  });

  test("5 text that does not fit its box is cut and reported with the slide number", () => {
    const long = "word ".repeat(400);
    const { bytes, warnings } = writePdf({ width: 400, height: 300, slides: [{ shapes: [slideBox("short", 300, 40)] }, { shapes: [slideBox(long, 300, 40)] }] });
    structure(bytes);
    expect(warnings).toEqual(["text was cut to fit its box on slide 2"]);
  });

  test("5 a table row taller than a page is cut and reported", () => {
    const cell = Array.from({ length: 200 }, (_, i) => para(`line ${i}`));
    const { warnings, bytes } = writePdf({ blocks: [{ type: "table", rows: [[cell]] }] });
    expect(warnings).toContain("a table row taller than a page was cut to fit");
    structure(bytes);
  });

  test("5 notes, and a table in a table, are reported as lost or flattened", () => {
    const inner = { type: "table" as const, rows: [[[para("x")], [para("y")]]] };
    const r = writePdf({ blocks: [{ type: "table", rows: [[[inner]]] }] });
    expect(r.warnings).toContain("a table inside a table was written as lines of text");
    expect(contentText(r.bytes)).toContain("x | y");
    const d = writePdf({ width: 100, height: 100, slides: [{ shapes: [], notes: "n" }] });
    expect(d.warnings).toContain("speaker notes are not written to a PDF");
  });

  test("6 the same model gives the same bytes, with no clock in the file", () => {
    const a = writePdf(textFixture).bytes;
    const b = writePdf(textFixture).bytes;
    expect(Buffer.from(a).equals(Buffer.from(b))).toBe(true);
    expect(latin(a)).toContain("/CreationDate (D:20000101000000Z)");
    expect(Buffer.from(writePdf(deckFixture).bytes).equals(Buffer.from(writePdf(deckFixture).bytes))).toBe(true);
  });

  test("7 a model that cannot be written fails with a FormatError", () => {
    const fails = (model: unknown, limits?: Parameters<typeof writePdf>[1]) => {
      try {
        writePdf(model as TextDoc, limits);
      } catch (e) {
        expect(e).toBeInstanceOf(FormatError);
        return (e as FormatError).code;
      }
      return "none";
    };
    expect(fails(null)).toBe("INVALID_MODEL");
    expect(fails({ blocks: "x" })).toBe("INVALID_MODEL");
    expect(fails({ blocks: [{ type: "nope" }] })).toBe("INVALID_MODEL");
    expect(fails({ blocks: [{ type: "heading", level: "1", content: [] }] })).toBe("INVALID_MODEL");
    expect(fails({ width: 10, height: 10, slides: [] })).toBe("INVALID_MODEL");
    expect(fails({ width: 10, height: 10, slides: [{ shapes: [{ type: "shape", kind: "line", frame: { x: 0, y: 0, width: NaN, height: 1 } }] }] })).toBe("INVALID_MODEL");
    let deep: unknown = { type: "paragraph", content: [] };
    for (let i = 0; i < 200; i++) deep = { type: "table", rows: [[[deep]]] };
    expect(fails({ blocks: [deep] })).toBe("INVALID_MODEL");
  });

  test("7 limits stop a big text, many pages, and a big output", () => {
    const fails = (model: TextDoc | Deck, limits: Parameters<typeof writePdf>[1]) => {
      try {
        writePdf(model, limits);
      } catch (e) {
        return (e as FormatError).code;
      }
      return "none";
    };
    expect(fails({ blocks: [para("x".repeat(100))] }, { maxTextChars: 50 })).toBe("LIMIT_EXCEEDED");
    expect(fails({ blocks: Array.from({ length: 300 }, () => para("p")) }, { maxEntries: 2 })).toBe("LIMIT_EXCEEDED");
    expect(fails(deckOf("a", "b", "c", "d", "e", "f", "g", "h", "i"), { maxEntries: 8 })).toBe("INVALID_MODEL");
    expect(fails({ blocks: [para("hello")] }, { maxTotalBytes: 200 })).toBe("LIMIT_EXCEEDED");
    const cells = Array.from({ length: 10 }, () => [para("c")]);
    expect(fails({ blocks: [{ type: "table", rows: [cells, cells] }] }, { maxCells: 5 })).toBe("LIMIT_EXCEEDED");
  });

  test("7 a footnote with an image or a footnote is refused, not a TypeError", () => {
    const note = (content: unknown) => ({ blocks: [{ type: "paragraph", content: [{ type: "footnote", content }] }] }) as unknown as TextDoc;
    expect(() => writePdf(note([{ type: "image", data: PNG, mime: "image/png" }]))).toThrow(FormatError);
    expect(() => writePdf(note([{ type: "footnote", content: [] }]))).toThrow(FormatError);
  });

  test("7 the text limit counts each character once and leaves image bytes out", () => {
    expect(() => writePdf({ blocks: [para("x".repeat(600))] }, { maxTextChars: 1000 })).not.toThrow();
    const big = new Uint8Array(PNG.length + 5000);
    big.set(PNG);
    const img = { type: "paragraph" as const, content: [{ type: "image" as const, data: big, mime: "image/png" as const, width: 10, height: 10 }] };
    expect(() => writePdf({ blocks: [img] }, { maxTextChars: 1000 })).not.toThrow();
  });

  test("7 PNGs with alpha share one decode budget, and equal bytes are stored once", () => {
    const copies = Array.from({ length: 6 }, () => ({ type: "paragraph" as const, content: [{ type: "image" as const, data: PNG.slice(), mime: "image/png" as const }] }));
    const { bytes } = writePdf({ blocks: copies });
    expect(latin(bytes).match(/\/Subtype \/Image/g)).toHaveLength(2);
    // The 1x1 fixture unpacks to 6 bytes: a budget of 6 bytes refuses it (as unpacked pixels or as output).
    const one = { blocks: copies.slice(0, 1) };
    expect(() => writePdf(one, { maxTotalBytes: 6 })).toThrow(/unpack|passes/);
    expect(() => writePdf(one, { maxTotalBytes: 100_000 })).not.toThrow();
  });

  test("4 a PNG with a tRNS chunk reports that its transparency is lost", () => {
    const withTrns = (() => {
      // Chunks are copied from the fixture; a tRNS chunk is put before IDAT. The reader of the PDF never checks its CRC.
      const at = PNG.findIndex((_, i) => i > 8 && String.fromCharCode(...PNG.subarray(i, i + 4)) === "IDAT") - 4;
      const trns = Uint8Array.of(0, 0, 0, 1, 0x74, 0x52, 0x4e, 0x53, 0, 0, 0, 0, 0);
      return new Uint8Array([...PNG.subarray(0, at), ...trns, ...PNG.subarray(at)]);
    })();
    // Colour type 6 has its own alpha, so only palette, gray, and RGB files report.
    const gray = withTrns.slice();
    gray[25] = 0;
    gray[24] = 8;
    const { warnings } = writePdf({ blocks: [{ type: "paragraph", content: [{ type: "image", data: gray, mime: "image/png" }] }] });
    expect(warnings).toContain("PNG transparency from a tRNS chunk is not written");
  });

  test("2 a list nested very deep stays on the page and says so", () => {
    let list: unknown = { type: "list", ordered: false, items: [{ content: [text("leaf")] }] };
    for (let i = 0; i < 30; i++) list = { type: "list", ordered: false, items: [{ content: [text(`l${i}`)], sublist: list }] };
    const { bytes, warnings } = writePdf({ blocks: [list as never] });
    expect(warnings).toContain("a list nested too deep was written at the deepest indent that fits");
    if (hasPdftotext) expect(pdftotext(bytes)).toContain("leaf");
  });

  test("2 spaces at the start of a paragraph are kept", () => {
    expect(contentText(writePdf({ blocks: [para("        indented")] }).bytes)).toBe("        indented");
  });

  test("7 a seeded mutation of a model gives a file or a FormatError, never another error", () => {
    let seed = 12345;
    const rnd = () => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) / 4294967296);
    const junk: unknown[] = [null, undefined, NaN, -1, 1e999, "", "x", {}, [], Symbol("s"), true, 0, 4294967296, { type: "text" }];
    const mutate = (v: unknown, depth = 0): unknown => {
      if (Array.isArray(v)) return v.map((x) => (rnd() < 0.15 ? junk[Math.floor(rnd() * junk.length)] : mutate(x, depth + 1)));
      if (v instanceof Uint8Array) return rnd() < 0.3 ? v.slice(0, Math.floor(rnd() * v.length)) : v;
      if (typeof v === "object" && v !== null) return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, rnd() < 0.1 ? junk[Math.floor(rnd() * junk.length)] : mutate(x, depth + 1)]));
      return v;
    };
    for (let i = 0; i < 400; i++) {
      for (const model of [textFixture, deckFixture]) {
        try {
          writePdf(mutate(model) as never);
        } catch (e) {
          expect(e).toBeInstanceOf(FormatError);
        }
      }
    }
  });

  test("4 a PNG bomb is stopped while it unpacks, and a bad filter byte is damage", () => {
    const png = (ctype: number, idat: Uint8Array, w = 1, h = 1) => {
      const chunk = (type: string, body: Uint8Array) => {
        const b = new Uint8Array(12 + body.length);
        new DataView(b.buffer).setUint32(0, body.length);
        b.set([...type].map((c) => c.charCodeAt(0)), 4);
        b.set(body, 8);
        return b;
      };
      const ihdr = new Uint8Array(13);
      const v = new DataView(ihdr.buffer);
      v.setUint32(0, w);
      v.setUint32(4, h);
      ihdr.set([8, ctype, 0, 0, 0], 8);
      return new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, ...chunk("IHDR", ihdr), ...chunk("IDAT", idat), ...chunk("IEND", new Uint8Array())]);
    };
    const inline = (data: Uint8Array) => ({ blocks: [{ type: "paragraph" as const, content: [{ type: "image" as const, data, mime: "image/png" as const }] }] });
    const bomb = png(6, zlibSync(new Uint8Array(64 * 1024 * 1024), { level: 9 }));
    const r = writePdf(inline(bomb));
    expect(r.warnings).toContain("an image is a PNG with damaged pixels: not written");
    const badFilter = writePdf(inline(png(6, zlibSync(Uint8Array.of(9, 0, 0, 0, 0)))));
    expect(badFilter.warnings).toContain("an image is a PNG with damaged pixels: not written");
  }, 60_000);

  test("4 a picture with no data is a FormatError, and a nested table reports lost images", () => {
    const deck = { width: 100, height: 100, slides: [{ shapes: [{ type: "picture", frame: { x: 0, y: 0, width: 5, height: 5 } }] }] };
    expect(() => writePdf(deck as never)).toThrow(FormatError);
    const inner = { type: "table" as const, rows: [[[{ type: "paragraph" as const, content: [{ type: "image" as const, data: PNG, mime: "image/png" as const }] }]]] };
    const r = writePdf({ blocks: [{ type: "table", rows: [[[inner]]] }] });
    expect(r.warnings).toContain("an image or footnote inside a nested table was not written");
  });

  test("7 long inputs do not overflow the stack: 150000 breaks, 150000 list items", () => {
    const run = Array.from({ length: 150_000 }, () => [text("a"), { type: "break" as const }]).flat();
    expect(() => writePdf({ blocks: [{ type: "paragraph", content: run }] }, { maxEntries: 20_000 })).not.toThrow();
    const items = Array.from({ length: 150_000 }, () => ({ content: [text("w")] }));
    expect(() => writePdf({ blocks: [{ type: "list", ordered: false, items }] }, { maxEntries: 20_000 })).not.toThrow();
  }, 60_000);

  test("7 a sublist shared by many items stops at the work cap, fast", () => {
    const leaf = { type: "list" as const, ordered: false, items: [{ content: [text("x")] }] };
    const level = (sub: unknown) => ({ type: "list" as const, ordered: false, items: Array.from({ length: 400 }, () => ({ content: [text("i")], sublist: sub as never })) });
    expect(() => writePdf({ blocks: [level(level(level(leaf)))] })).toThrow(FormatError);
  });

  test("7 footnotes nested 200000 deep are refused without recursion", () => {
    let note: unknown = { type: "footnote", content: [text("end")] };
    for (let i = 0; i < 200_000; i++) note = { type: "footnote", content: [note] };
    expect(() => writePdf({ blocks: [{ type: "paragraph", content: [note as never] }] })).toThrow(FormatError);
  });

  test("2 a change of style inside a word is not a place to break", () => {
    const { bytes } = writePdf({ blocks: [{ type: "paragraph", content: [text(`${"b".repeat(40)} ${"x".repeat(40)}`), { type: "text", text: "cccccc", marks: ["bold"] }, text(", after")] }] });
    const lines = pdftotextOrNull(bytes);
    if (lines) expect(lines).toMatch(/x+c{6},/);
  });

  test("4 a small PNG that declares a huge size is refused, not passed to the viewer", () => {
    const huge = PNG.slice();
    huge[25] = 2;
    new DataView(huge.buffer).setUint32(16, 60_000);
    new DataView(huge.buffer).setUint32(20, 60_000);
    expect(() => writePdf({ blocks: [{ type: "paragraph", content: [{ type: "image", data: huge, mime: "image/png" }] }] })).toThrow(/unpacked|unpack/);
  });
});
const pdftotextOrNull = (bytes: Uint8Array): string | null => (hasPdftotext ? pdftotext(bytes) : null);

describe("PDF writer, review round 4", () => {
  test("7 a table of empty rows under a wide row is refused fast", () => {
    const wide = Array.from({ length: 2000 }, () => [] as never[]);
    const empty: never[][] = [];
    expect(() => writePdf({ blocks: [{ type: "table", rows: [wide, ...Array.from({ length: 1000 }, () => empty)] }] }, { maxCells: 100_000 })).toThrow(FormatError);
  });

  test("4 a plain PNG with damaged pixels or a bomb in its data is reported", () => {
    const rebuilt = (idat: Uint8Array) => {
      const at = PNG.findIndex((_, i) => i > 8 && String.fromCharCode(...PNG.subarray(i, i + 4)) === "IDAT") - 4;
      const head = PNG.slice(0, at);
      head[25] = 2;
      const chunk = new Uint8Array(12 + idat.length);
      new DataView(chunk.buffer).setUint32(0, idat.length);
      chunk.set([0x49, 0x44, 0x41, 0x54], 4);
      chunk.set(idat, 8);
      return new Uint8Array([...head, ...chunk]);
    };
    const doc = (data: Uint8Array) => ({ blocks: [{ type: "paragraph" as const, content: [{ type: "image" as const, data, mime: "image/png" as const }] }] });
    const damaged = "an image is a PNG with damaged pixels: not written";
    expect(writePdf(doc(rebuilt(Uint8Array.of(1, 2, 3, 4, 5)))).warnings).toContain(damaged);
    expect(writePdf(doc(rebuilt(zlibSync(new Uint8Array(32 * 1024 * 1024))))).warnings).toContain(damaged);
    expect(writePdf(doc(rebuilt(zlibSync(Uint8Array.of(0, 1, 2, 3))))).warnings).toEqual([]);
  });

  test("5 image descriptions, bold code, and a long title are reported", () => {
    const r = writePdf({ title: "t".repeat(600), blocks: [{ type: "paragraph", content: [{ type: "text", text: "c", marks: ["code", "bold"] }, { type: "image", data: PNG, mime: "image/png", alt: "dot" }] }] });
    expect(r.warnings).toEqual(expect.arrayContaining(["image descriptions are not written to a PDF", "bold and italic on code text are written as plain Courier", "the title was cut to 500 characters"]));
  });

  test("2 spaces after a hard break are kept, and equal images give one resource name", () => {
    const run = [text("def f():"), { type: "break" as const }, text("    return 1")];
    expect(contentText(writePdf({ blocks: [{ type: "paragraph", content: run }] }).bytes)).toBe("def f():    return 1");
    const img = () => ({ type: "paragraph" as const, content: [{ type: "image" as const, data: PNG.slice(), mime: "image/png" as const }] });
    const names = /\/XObject <<([^>]*)>>/.exec(latin(writePdf({ blocks: [img(), img(), img()] }).bytes))![1].match(/\/Im\d+/g)!;
    expect(new Set(names).size).toBe(names.length);
  });

  test("5 an empty paragraph in a short box is not reported as cut text", () => {
    const deck = { width: 100, height: 100, slides: [{ shapes: [{ type: "textbox" as const, frame: { x: 0, y: 0, width: 50, height: 2 }, paragraphs: [[]] }] }] };
    expect(writePdf(deck).warnings).toEqual([]);
  });

  test("4 palette and sub-byte gray PNGs are embedded", () => {
    const png = (ctype: number, depth: number, w: number, h: number, rows: number[], plte?: number[]) => {
      const chunk = (type: string, body: Uint8Array) => {
        const b = new Uint8Array(12 + body.length);
        new DataView(b.buffer).setUint32(0, body.length);
        b.set([...type].map((c) => c.charCodeAt(0)), 4);
        b.set(body, 8);
        return b;
      };
      const ihdr = new Uint8Array(13);
      new DataView(ihdr.buffer).setUint32(0, w);
      new DataView(ihdr.buffer).setUint32(4, h);
      ihdr.set([depth, ctype, 0, 0, 0], 8);
      return new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, ...chunk("IHDR", ihdr), ...(plte ? chunk("PLTE", Uint8Array.from(plte)) : []), ...chunk("IDAT", zlibSync(Uint8Array.from(rows))), ...chunk("IEND", new Uint8Array())]);
    };
    const doc = (data: Uint8Array) => ({ blocks: [{ type: "paragraph" as const, content: [{ type: "image" as const, data, mime: "image/png" as const }] }] });
    const palette = writePdf(doc(png(3, 8, 2, 2, [0, 0, 1, 0, 1, 0], [255, 0, 0, 0, 0, 255])));
    expect(palette.warnings).toEqual([]);
    expect(latin(palette.bytes)).toContain("/Indexed /DeviceRGB 1");
    const bit = writePdf(doc(png(0, 1, 16, 1, [0, 0xaa, 0x55])));
    expect(bit.warnings).toEqual([]);
    expect(latin(bit.bytes)).toContain("/BitsPerComponent 1");
  });

  test("7 many small styled runs are stopped by the output cap", () => {
    const runs = Array.from({ length: 20_000 }, (_, i) => ({ type: "text" as const, text: "ab"[i % 2], marks: [(["bold", "underline", "strike"] as const)[i % 3]] }));
    expect(() => writePdf({ blocks: [{ type: "paragraph", content: runs }] }, { maxTotalBytes: 100_000 })).toThrow(FormatError);
  });

  test("7 a table of very narrow columns lays out a page of lines for each cell, not a line for each character", () => {
    const row = Array.from({ length: 60 }, () => [para("w".repeat(5000))]);
    const { warnings, bytes } = writePdf({ blocks: [{ type: "table", rows: Array.from({ length: 5 }, () => row) }] });
    structure(bytes);
    expect(warnings).toContain("a table row taller than a page was cut to fit");
  });

  test("2 wide Latin-1 letters wrap inside the page", () => {
    const { bytes } = writePdf({ blocks: [para("ÆŒ".repeat(60))] });
    // Each of these letters is 11 points wide at 11 point type, so a 468 point line holds 42 at most.
    const lines = [...latin(bytes).matchAll(/\/Filter \/FlateDecode \/Length (\d+)>>\nstream\n/g)].flatMap((m) => {
      const at = m.index! + m[0].length;
      return [...strFromU8(unzlibSync(bytes.subarray(at, at + Number(m[1]))), true).matchAll(/\(((?:[^\\)]|\\.)*)\) Tj/g)].map((t) => t[1].replace(/\\\d{3}/g, "?").length / 4 === 0 ? 0 : (t[1].match(/\\\d{3}/g) ?? []).length);
    });
    expect(lines.length).toBeGreaterThan(2);
    for (const n of lines) expect(n).toBeLessThanOrEqual(42);
  });

  test("7 a zero-width slide box with millions of characters is cut, not laid out line by line", () => {
    const deck = { width: 100, height: 100, slides: [{ shapes: [{ type: "textbox" as const, frame: { x: 0, y: 0, width: 0, height: 10 }, paragraphs: [[text("a".repeat(4_000_000))]] }] }] };
    const { warnings, bytes } = writePdf(deck);
    structure(bytes);
    expect(warnings).toContain("text was cut to fit its box on slide 1");
  }, 60_000);

  test("7 a title is counted against the text limit", () => {
    expect(() => writePdf({ title: "x".repeat(2000), blocks: [] }, { maxTextChars: 1000 })).toThrow(FormatError);
    expect(writePdf({ title: "x".repeat(900), blocks: [] }).warnings).toContain("the title was cut to 500 characters");
  });

  test("5 text lost behind a run of line breaks is reported, and a soft hyphen is dropped", () => {
    const runs = [text("a"), ...Array.from({ length: 10 }, () => text("\n")), text("SECRET")];
    const r = writePdf({ width: 300, height: 100, slides: [{ shapes: [{ type: "textbox", frame: { x: 0, y: 0, width: 300, height: 30 }, paragraphs: [runs] }] }] });
    expect(r.warnings).toContain("text was cut to fit its box on slide 1");
    expect(contentText(r.bytes)).not.toContain("SECRET");
    expect(contentText(writePdf({ blocks: [para("co\u00adop")] }).bytes)).toBe("coop");
  });

  test("2 a line break inside a run is a line break; a long numbered list keeps its numbers clear of the text", () => {
    const deck = { width: 300, height: 100, slides: [{ shapes: [{ type: "textbox" as const, frame: { x: 0, y: 0, width: 300, height: 60 }, paragraphs: [[text("Line one\nLine two")]] }] }] };
    expect(contentText(writePdf(deck).bytes)).toBe("Line oneLine two");
    expect(latin(writePdf(deck).bytes)).not.toContain("Line one Line two");
    if (hasPdftotext) expect(pdftotext(writePdf(deck).bytes)).toMatch(/Line one\s*\n\s*Line two/);
    const items = Array.from({ length: 120 }, (_, i) => ({ content: [text(`item ${i + 1}`)] }));
    if (hasPdftotext) expect(pdftotext(writePdf({ blocks: [{ type: "list", ordered: true, items }] }).bytes)).toMatch(/100\.\s+item 100/);
  });

  test("5 an image that fills a table cell does not report a cut row", () => {
    const tall = { type: "paragraph" as const, content: [{ type: "image" as const, data: PNG, mime: "image/png" as const, width: 300, height: 2000 }] };
    expect(writePdf({ blocks: [{ type: "table", rows: [[[tall]]] }] }).warnings).toEqual([]);
  });

  test("5 blocks dropped behind blank space in a cell are reported", () => {
    const cell = [...Array.from({ length: 30 }, (_, i) => para(`row ${i}`)), ...Array.from({ length: 8 }, () => ({ type: "list" as const, ordered: false, items: [] })), para("LOSTWORD")];
    const { warnings, bytes } = writePdf({ blocks: [{ type: "table", rows: [[cell]] }] });
    expect(contentText(bytes)).not.toContain("LOSTWORD");
    expect(warnings).toContain("a table row taller than a page was cut to fit");
  });

  test("2 a lone surrogate in a title is replaced", () => {
    expect(latin(writePdf({ title: "a\ud800b", blocks: [] }).bytes)).toContain("/Title <FEFF0061fffd0062>");
  });
});

