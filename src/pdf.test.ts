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
 *  7 a writer takes a model it cannot write: no slides, nesting too deep, a bad block; a limit is missing
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
  expect(s.startsWith("%PDF-1.4\n")).toBe(true);
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
});
