/*
 * DOCX, end to end: the real writer, the real reader, an independent XML parser, and LibreOffice.
 * The ways it can fail are written first; each has a case below.
 *
 *  1 text is lost or changed: marks, links, runs of spaces, tabs, line breaks, markup characters
 *  2 structure is lost: headings, nested lists, two lists in a row that merge or keep counting,
 *    tables with a header and empty cells
 *  3 an image or a footnote is lost, or an image takes a hostile size
 *  4 the package is malformed: a part does not parse, a part is missing from the content types,
 *    a relationship points nowhere
 *  5 the same model gives different bytes
 *  6 an unsafe link or a bad image drops without a word; a size limit is missing
 *  7 a hostile file is read: a DOCTYPE, the wrong type, a remote or external image, a script link,
 *    XML nested without end, a footnote that is not there, tracked deletions read as text
 *  8 LibreOffice cannot open our file, or its re-save does not read back to the same model, in the
 *    same format or after a trip through ODT
 */
import { isDeepStrictEqual } from "node:util";
import { DOMParser } from "@xmldom/xmldom";
import { strFromU8, unzipSync } from "fflate";
import { afterAll, describe, expect, test } from "vitest";
import { readDocx, writeDocx } from "./docx.js";
import { FormatError, withLimits } from "./limits.js";
import { normalizeTextDoc, type Block, type List, type TextDoc } from "./model.js";
import { readOdt, writeOdt } from "./odt.js";
import { PNG, textFixture as fixture } from "./testing/fixtures.js";
import { LibreOffice, hasLibreOffice, libreOfficeRequired, sha256, writeArtifact, type RoundTripRecord } from "./testing/libreoffice.js";
import { readXlsx, writeXlsx } from "./xlsx.js";
import { writeZip } from "./zip.js";

const text = (t: string) => ({ type: "text" as const, text: t });
const para = (t: string): Block => ({ type: "paragraph", content: [text(t)] });
const files = (bytes: Uint8Array) => unzipSync(bytes);
const patch = (bytes: Uint8Array, name: string, change: (xml: string) => string): Uint8Array =>
  writeZip(Object.entries(files(bytes)).map(([n, data]) => ({ name: n, data: n === name ? change(strFromU8(data)) : data })), withLimits());
const wellFormed = (xml: string): void => {
  const seen: string[] = [];
  new DOMParser({ onError: (level, message) => seen.push(`${level}: ${message}`) }).parseFromString(xml, "text/xml");
  expect(seen).toEqual([]);
};

describe("write then read", () => {
  const written = writeDocx(fixture);

  test("1-3 the model reads back unchanged, with no warnings", () => {
    const read = readDocx(written.bytes);
    expect(read.model).toEqual(normalizeTextDoc(fixture));
    expect(read.warnings).toEqual([]);
    expect(written.warnings).toEqual([]);
  });

  test("2 lists in a row stay apart, and an ordered list restarts at one", () => {
    const list = (ordered: boolean, ...items: string[]): Block => ({ type: "list", ordered, items: items.map((i) => ({ content: [text(i)] })) });
    const doc: TextDoc = { blocks: [list(true, "a", "b"), list(true, "c"), list(false, "d"), list(false, "e"), para("between"), list(true, "f")] };
    expect(readDocx(writeDocx(doc).bytes).model).toEqual(doc);
    const numbering = strFromU8(files(writeDocx(doc).bytes)["word/numbering.xml"]);
    expect(numbering.match(/<w:startOverride w:val="1"\/>/g)).toHaveLength(9 * 5);
  });

  test("2 a list nests across kinds and levels, and a list item may be empty", () => {
    const doc: TextDoc = {
      blocks: [
        {
          type: "list",
          ordered: true,
          items: [
            { content: [text("one")], sublist: { type: "list", ordered: false, items: [{ content: [text("bullet")], sublist: { type: "list", ordered: true, items: [{ content: [text("deep")] }] } }, { content: [] }] } },
            { content: [text("two")] },
          ],
        },
      ],
    };
    expect(readDocx(writeDocx(doc).bytes).model).toEqual(doc);
  });

  test("4 every XML part parses with an independent parser", () => {
    const all = files(written.bytes);
    const xml = Object.keys(all).filter((n) => /\.(xml|rels)$/.test(n));
    expect(xml.length).toBeGreaterThanOrEqual(8);
    for (const name of xml) wellFormed(strFromU8(all[name]));
  });

  test("4 every part is in the content types, and every internal relationship target exists", () => {
    const all = files(written.bytes);
    const types = strFromU8(all["[Content_Types].xml"]);
    for (const part of Object.keys(all).filter((n) => n.endsWith(".xml") && n !== "[Content_Types].xml")) expect(types).toContain(`PartName="/${part}"`);
    expect(Object.keys(all)[0]).toBe("[Content_Types].xml");
    const rels = strFromU8(all["word/_rels/document.xml.rels"]);
    for (const m of rels.matchAll(/Target="([^"]+)"(?! TargetMode)/g)) expect(all[`word/${m[1]}`]).toBeDefined();
    expect(rels.match(/TargetMode="External"/g)).toHaveLength(2);
  });

  test("5 the same model gives the same bytes", () => {
    expect(writeDocx(fixture).bytes).toEqual(written.bytes);
  });
});

describe("write guards", () => {
  test("3 an image over the entry limit, and an export over the total limit, are refused", () => {
    const big = new Uint8Array(PNG.length + 4000);
    big.set(PNG);
    const doc: TextDoc = { blocks: [{ type: "paragraph", content: [{ type: "image", data: big, mime: "image/png" }] }] };
    expect(() => writeDocx(doc, { maxEntryBytes: 1000 })).toThrow(FormatError);
    expect(() => writeDocx(doc, { maxTotalBytes: 6000 })).toThrow(/export is larger/);
  });

  test("3 a hostile image size is clamped to 22 inches in EMU", () => {
    const doc: TextDoc = { blocks: [{ type: "paragraph", content: [{ type: "image", data: PNG, mime: "image/png", width: 1e12, height: 5e11 }] }] };
    const xml = strFromU8(files(writeDocx(doc).bytes)["word/document.xml"]);
    expect(Number(/<wp:extent cx="(\d+)"/.exec(xml)![1])).toBeLessThanOrEqual(1584 * 12700);
  });

  test("6 an unsafe link and a bad image are reported", () => {
    const doc: TextDoc = {
      blocks: [{ type: "paragraph", content: [{ type: "text", text: "click", href: "javascript:alert(1)" }, { type: "image", data: new TextEncoder().encode("<svg onload=alert(1)/>"), mime: "image/png" }] }],
    };
    const { bytes, warnings } = writeDocx(doc);
    expect(warnings).toEqual(["a link with an unsafe target was written as plain text", "an image is not a PNG, JPEG, or GIF: not written"]);
    const all = files(bytes);
    expect(strFromU8(all["word/document.xml"])).toContain("click");
    expect(strFromU8(all["word/_rels/document.xml.rels"])).not.toMatch(/javascript/);
    expect(Object.keys(all).some((n) => n.startsWith("word/media"))).toBe(false);
  });
});

describe("hostile files", () => {
  const base = writeDocx({ blocks: [para("x")] }).bytes;
  const edit = (change: (xml: string) => string, extra?: (bytes: Uint8Array) => Uint8Array) => (extra ?? ((b) => b))(patch(base, "word/document.xml", change));

  test("7 a DOCTYPE with an entity is refused", () => {
    expect(() => readDocx(edit((x) => x.replace("?>", `?><!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]>`)))).toThrow(/DOCTYPE/);
  });

  test("7 the wrong type of file is refused", () => {
    expect(() => readDocx(writeXlsx({ sheets: [{ name: "S", cells: {} }] }).bytes)).toThrow(/not an DOCX/);
    expect(() => readDocx(writeOdt({ blocks: [] }).bytes)).toThrow(FormatError);
    expect(() => readDocx(new TextEncoder().encode("PK nope"))).toThrow(FormatError);
    expect(() => readXlsx(base)).toThrow(FormatError);
  });

  test("7 a remote image, an external image relationship, and a script link are not followed", () => {
    const image = (attrs: string) => `<w:r><w:drawing><wp:inline><wp:extent cx="1" cy="1"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip ${attrs}/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const link = `<w:hyperlink r:id="rId50"><w:r><w:t>this</w:t></w:r></w:hyperlink>`;
    const bytes = patch(
      patch(base, "word/document.xml", (x) => x.replace("</w:p>", `${image('r:link="rId60"')}${image('r:embed="rId61"')}${link}</w:p>`)),
      "word/_rels/document.xml.rels",
      (x) =>
        x.replace(
          "</Relationships>",
          `<Relationship Id="rId50" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="javascript:alert(1)" TargetMode="External"/><Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="https://evil.example/p.png" TargetMode="External"/></Relationships>`,
        ),
    );
    const read = readDocx(bytes);
    expect(read.model.blocks).toEqual([{ type: "paragraph", content: [text("xthis")] }]);
    expect(read.warnings.sort()).toEqual(["a link with an unsafe target was read as plain text", "a remote image: not loaded", "an image that is not a PNG, JPEG, or GIF: not imported"]);
  });

  test("7 an image relationship that leaves the package is not read", () => {
    const withImage = writeDocx({ blocks: [{ type: "paragraph", content: [{ type: "image", data: PNG, mime: "image/png" }] }] }).bytes;
    const read = readDocx(patch(withImage, "word/_rels/document.xml.rels", (x) => x.replace("media/image1.png", "../../../etc/passwd")));
    expect(read.model.blocks).toEqual([{ type: "paragraph", content: [] }]);
    expect(read.warnings).toEqual(["an image that is not a PNG, JPEG, or GIF: not imported"]);
  });

  test("7 tables nested past the XML depth cap are refused", () => {
    const deep = "<w:tbl><w:tr><w:tc>".repeat(60) + "<w:p/>" + "</w:tc></w:tr></w:tbl>".repeat(60);
    expect(() => readDocx(edit((x) => x.replace("</w:p>", `</w:p>${deep}`)))).toThrow(/nests deeper/);
  });

  test("7 a missing footnote warns, tracked insertions read, tracked deletions are skipped", () => {
    const read = readDocx(
      edit((x) => x.replace("</w:p>", `<w:r><w:footnoteReference w:id="9"/></w:r><w:ins><w:r><w:t> new</w:t></w:r></w:ins><w:del><w:r><w:delText> gone</w:delText></w:r></w:del></w:p>`)),
    );
    expect(read.model.blocks).toEqual([{ type: "paragraph", content: [text("x new")] }]);
    expect(read.warnings.sort()).toEqual(["a footnote that is missing: not imported", "tracked deletions: skipped"]);
  });
});

describe("a link target that many runs share", () => {
  test("12 each run that names a shared target counts it, so a small file cannot make a model of gigabytes", { timeout: 60_000 }, () => {
    const base = writeDocx({ blocks: [{ type: "paragraph", content: [{ type: "text", text: "a", href: "https://example.org/x" }, text("b")] }] }).bytes;
    const target = `https://example.org/${"u".repeat(100_000)}`;
    const pair = `<w:hyperlink r:id="rId4"><w:r><w:t>a</w:t></w:r></w:hyperlink><w:r><w:t>b</w:t></w:r>`;
    const make = (pairs: number) => patch(patch(base, "word/_rels/document.xml.rels", (x) => x.replace(/Target="https:\/\/example\.org\/x"/, `Target="${target}"`)), "word/document.xml", (x) => x.replace(/<w:p>.*<\/w:p>/, `<w:p>${pair.repeat(pairs)}</w:p>`));
    expect(readDocx(make(50)).model.blocks).toHaveLength(1);
    const started = Date.now();
    expect(() => readDocx(make(200))).toThrow(/characters of run text and link targets/);
    expect(Date.now() - started).toBeLessThan(5000);
  });
});

describe("a link inside a footnote", () => {
  const doc: TextDoc = { blocks: [{ type: "paragraph", content: [text("claim"), { type: "footnote", content: [text("see "), { type: "text", text: "source", href: "https://example.org/n" }] }] }] };

  test("11 the link is written to the footnotes part, which has relationships of its own, and reads back", () => {
    const all = files(writeDocx(doc).bytes);
    expect(strFromU8(all["word/_rels/footnotes.xml.rels"])).toContain('Target="https://example.org/n"');
    wellFormed(strFromU8(all["word/_rels/footnotes.xml.rels"]));
    expect(strFromU8(all["word/_rels/document.xml.rels"])).not.toContain("example.org/n");
    const read = readDocx(writeDocx(doc).bytes);
    expect(read.model).toEqual(normalizeTextDoc(doc));
    expect(read.warnings).toEqual([]);
  });

  test("11 a link id in a footnote is not read from the document's relationships, and only a hyperlink is a link", () => {
    const withoutRels = Object.entries(files(writeDocx(doc).bytes))
      .filter(([name]) => name !== "word/_rels/footnotes.xml.rels")
      .map(([name, data]) => ({ name, data }));
    const read = readDocx(writeZip(withoutRels, withLimits()));
    expect(JSON.stringify(read.model)).not.toMatch(/styles\.xml|example\.org/);
    expect(read.warnings).toEqual(["a link whose relationship is missing or is not a link was read as plain text"]);
    // The id names the document's styles relationship, not a link.
    const styles = patch(writeDocx({ blocks: [para("x")] }).bytes, "word/document.xml", (x) => x.replace("<w:r>", `<w:hyperlink r:id="rId1"><w:r><w:t>y</w:t></w:r></w:hyperlink><w:r>`));
    const toStyles = readDocx(styles);
    expect(JSON.stringify(toStyles.model)).not.toContain("styles.xml");
    expect(toStyles.warnings).toEqual(["a link whose relationship is missing or is not a link was read as plain text"]);
  });

  test("11 many footnotes and a large footnotes relationships part read in linear time", { timeout: 60_000 }, () => {
    const base = writeDocx(doc).bytes;
    const links = Array.from({ length: 9000 }, (_, i) => `<Relationship Id="rIdX${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.org/${i}" TargetMode="External"/>`).join("");
    const notes = Array.from({ length: 200_000 }, (_, i) => `<w:footnote w:id="${i + 10}"/>`).join("");
    const big = patch(patch(base, "word/_rels/footnotes.xml.rels", (x) => x.replace("</Relationships>", `${links}</Relationships>`)), "word/footnotes.xml", (x) => x.replace("</w:footnotes>", `${notes}</w:footnotes>`));
    const started = Date.now();
    expect(readDocx(big).model.blocks).toHaveLength(1);
    expect(Date.now() - started).toBeLessThan(10_000);
  });

  test("11 a footnote that holds a footnote, an image, or itself is refused, not followed", () => {
    const cyclic = { type: "footnote" as const, content: [] as unknown[] };
    cyclic.content.push(cyclic);
    const nested = { type: "footnote" as const, content: [{ type: "footnote", content: [text("x")] }] };
    const withImage = { type: "footnote" as const, content: [{ type: "image", data: PNG, mime: "image/png" }] };
    for (const note of [cyclic, nested, withImage]) {
      expect(() => writeDocx({ blocks: [{ type: "paragraph", content: [note as never] }] })).toThrow(/a footnote holds only text and line breaks/);
    }
    let deep: unknown = { type: "footnote", content: [text("x")] };
    for (let i = 0; i < 20_000; i++) deep = { type: "footnote", content: [deep] };
    expect(() => writeDocx({ blocks: [{ type: "paragraph", content: [deep as never] }] })).toThrow(FormatError);
  });

  test("11 a model of the wrong shape is an INVALID_MODEL error, not a TypeError", () => {
    const wrong = (blocks: unknown) => ({ blocks }) as unknown as TextDoc;
    for (const model of [wrong(null), wrong([null]), wrong([{ type: "list", ordered: false, items: [null] }]), wrong([{ type: "table", rows: [[["x"]]] }]), wrong([{ type: "paragraph", content: [{ type: "text", text: "x", marks: 5 }] }])]) {
      expect(() => writeDocx(model), JSON.stringify(model)).toThrow(FormatError);
    }
  });

  test("11 an image or a note inside a footnote that is dropped on read is reported", () => {
    const two: TextDoc = { blocks: [{ type: "paragraph", content: [{ type: "footnote", content: [text("one")] }, { type: "footnote", content: [text("two")] }] }] };
    const base = writeDocx(two).bytes;
    // The second note holds a reference to the first, which the model cannot hold.
    const read = readDocx(patch(base, "word/footnotes.xml", (x) => x.replace(/(<w:footnote w:id="2">.*?)<\/w:p>/, `$1<w:r><w:footnoteReference w:id="1"/></w:r></w:p>`)));
    expect(read.warnings).toContain("an image or a note inside a footnote: not imported");
  });
});

describe("a footnote that many references point to", () => {
  const base = writeDocx({ blocks: [para("x")] }).bytes;
  const runs = Array.from({ length: 2000 }, (_, i) => `<w:r>${i % 2 ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t>n${i}</w:t></w:r>`).join("");
  const note = `<w:footnote w:id="1"><w:p>${runs}</w:p></w:footnote>`;
  const reference = `<w:r><w:footnoteReference w:id="1"/></w:r>`;
  const hostile = (references: number) =>
    patch(
      patch(base, "word/document.xml", (x) => x.replace("</w:p>", `${reference.repeat(references)}</w:p>`)),
      "word/footnotes.xml",
      (x) => x.replace("</w:footnotes>", `${note}</w:footnotes>`),
    );

  test("9 a reader keeps the note once, warns on a repeat, and stays fast", () => {
    const started = Date.now();
    const read = readDocx(hostile(2000));
    expect(Date.now() - started).toBeLessThan(1500);
    const paragraph = read.model.blocks[0];
    if (paragraph.type !== "paragraph") throw new Error("expected a paragraph");
    expect(paragraph.content.filter((n) => n.type === "footnote")).toHaveLength(1);
    expect(read.warnings).toEqual(["a repeated footnote reference: not imported"]);
  });

  test("9 a note used once reads with no warning", () => {
    const read = readDocx(hostile(1));
    expect(read.warnings).toEqual([]);
  });

  test("9 a writer refuses a model that holds one footnote many times past the text cap", () => {
    const shared = { type: "footnote" as const, content: [text("n".repeat(1000))] };
    const doc: TextDoc = { blocks: [{ type: "paragraph", content: Array.from({ length: 50 }, () => shared) }] };
    expect(() => writeDocx(doc, { maxTextChars: 10_000 })).toThrow(/more than 10000 characters of text/);
    expect(() => writeOdt(doc, { maxTextChars: 10_000 })).toThrow(/more than 10000 characters of text/);
    expect(writeDocx(doc).warnings).toEqual([]);
  });
});

describe("a writer that gets one object in many places", () => {
  const overCap = (doc: TextDoc, limits = { maxTotalBytes: 2_000_000 }, message = /while it is written|characters of text/) => {
    const started = Date.now();
    try {
      writeDocx(doc, limits);
    } catch (error) {
      expect(error).toBeInstanceOf(FormatError);
      expect((error as FormatError).code).toBe("LIMIT_EXCEEDED");
      expect((error as FormatError).message).toMatch(message);
      expect(Date.now() - started).toBeLessThan(5000);
      return;
    }
    throw new Error("expected LIMIT_EXCEEDED");
  };

  test("10 a note of breaks, empty runs, and a rows of empty cells stop at the byte cap", () => {
    const note = { type: "footnote" as const, content: Array.from({ length: 10_000 }, () => ({ type: "break" as const })) };
    overCap({ blocks: [{ type: "paragraph", content: Array.from({ length: 10_000 }, () => note) }] });
    const emptyRuns = { type: "footnote" as const, content: Array.from({ length: 10_000 }, () => ({ type: "text" as const, text: "" })) };
    overCap({ blocks: [{ type: "paragraph", content: Array.from({ length: 10_000 }, () => emptyRuns) }] });
    const row = Array.from({ length: 2000 }, () => [] as Block[]);
    overCap({ blocks: [{ type: "table", rows: Array.from({ length: 5000 }, () => row) }] });
  });

  test("10 nested tables and lists that hold the level below twice stop at the byte cap", () => {
    const table = (k: number): Block => {
      if (k === 0) return para("x");
      const inner = table(k - 1);
      return { type: "table", rows: [[[inner], [inner]]] };
    };
    overCap({ blocks: [table(40)] });
    const list = (k: number): Block => {
      const inner = k === 0 ? undefined : (list(k - 1) as List);
      return { type: "list", ordered: false, items: [{ content: [], ...(inner ? { sublist: inner } : {}) }, { content: [], ...(inner ? { sublist: inner } : {}) }] };
    };
    overCap({ blocks: [list(40)] });
  });

  test("10 a link target and an image description count when one run or image is used many times", () => {
    const link = { type: "text" as const, text: "x", href: `https://example.org/${"a".repeat(100_000)}` };
    overCap({ blocks: [{ type: "paragraph", content: Array.from({ length: 1000 }, () => link) }] }, { maxTotalBytes: 100_000_000, maxTextChars: 5_000_000 } as never);
    const image = { type: "image" as const, data: PNG, mime: "image/png" as const, alt: "d".repeat(100_000) };
    overCap({ blocks: [{ type: "paragraph", content: Array.from({ length: 1000 }, () => image) }] });
  });

  test("10 one image part named by many drawings counts its bytes for each drawing", () => {
    const padded = new Uint8Array(10_000);
    padded.set(PNG);
    const drawings = (n: number) => ({ blocks: [{ type: "paragraph" as const, content: Array.from({ length: n }, () => ({ type: "image" as const, data: padded, mime: "image/png" as const })) }] });
    const one = writeDocx(drawings(1)).bytes;
    const many = patch(one, "word/document.xml", (x) => x.replace(/<w:p>(.*)<\/w:p>/, (_m, inner: string) => `<w:p>${inner.repeat(5)}</w:p>`));
    expect(readDocx(many, { maxTotalBytes: 60_000 }).model.blocks).toHaveLength(1);
    expect(() => readDocx(many, { maxTotalBytes: 40_000 })).toThrow(/images in the file add up to more than 40000 bytes/);
  });

  test("10 a list level that is not one digit from 0 to 8 reads as level 0, with a warning", () => {
    const base = writeDocx({ blocks: [{ type: "list", ordered: false, items: [{ content: [text("a")] }] }] }).bytes;
    for (const level of ["1e9", "-5", "abc", "01", " 1", "0x1", "9"]) {
      const read = readDocx(patch(base, "word/document.xml", (x) => x.replace('<w:ilvl w:val="0"/>', `<w:ilvl w:val="${level}"/>`)));
      expect(read.model.blocks, level).toHaveLength(1);
      expect(read.warnings, level).toEqual(["a list level that is not 0 to 8 was read as level 0"]);
    }
  });

  test("10 empty lists charge the numbering they add", () => {
    const empty = { type: "list" as const, ordered: false, items: [] };
    overCap({ blocks: Array.from({ length: 100_000 }, () => empty) }, { maxTotalBytes: 1_000_000 }, /passes 1000000 bytes while it is written/);
  });

  test("10 a run with a text or link that is not a string, and a deep or cyclic model, are refused with a FormatError", () => {
    const doc = (run: unknown) => ({ blocks: [{ type: "paragraph", content: [run] }] }) as unknown as TextDoc;
    for (const run of [{ type: "text", text: Array(100).fill("a".repeat(1000)) }, { type: "text", text: 7 }, { type: "text", text: "x", href: { length: -1e15 } }]) {
      expect(() => writeDocx(doc(run)), JSON.stringify(run)).toThrow(/needs text and a link that are strings/);
    }
    expect(() => writeDocx(doc(null))).toThrow(/is not an object/);
    expect(() => writeDocx({ blocks: [{ type: "list", ordered: false }] } as unknown as TextDoc)).toThrow(/the items of a list must be a list/);
    expect(() => writeDocx({ blocks: [{ type: "paragraph", content: [{ type: "image", data: "x", mime: "image/png" }] }] } as unknown as TextDoc)).toThrow(/Uint8Array/);
    let deep: Block = { type: "list", ordered: false, items: [{ content: [] }] };
    for (let i = 0; i < 20_000; i++) deep = { type: "list", ordered: false, items: [{ content: [], sublist: deep as List }] };
    overCap({ blocks: [deep] }, undefined, /nests lists and tables deeper/);
    const cyclic = { type: "table" as const, rows: [[[] as Block[]]] };
    cyclic.rows[0][0].push(cyclic);
    overCap({ blocks: [cyclic] }, undefined, /nests lists and tables deeper/);
  });

  test("10 one run of eight million newlines is refused before its XML is made", () => {
    overCap({ blocks: [{ type: "paragraph", content: [{ type: "text", text: "\n".repeat(8_000_000) }] }] }, { maxTotalBytes: 100_000_000 }, /while it is written/);
  });

  test("10 a model node of an unknown kind is an INVALID_MODEL error, not a TypeError", () => {
    for (const doc of [{ blocks: [{ type: "foo" }] }, { blocks: [{ type: "paragraph", content: [{ type: "bar" }] }] }]) {
      expect(() => writeDocx(doc as unknown as TextDoc)).toThrow(/is not a kind of/);
    }
  });

  test("10 a body, a content control, and a table with 200,000 blocks or rows fail or pass as a FormatError, never a stack overflow", { timeout: 60_000 }, () => {
    const base = writeDocx({ blocks: [para("x")] }).bytes;
    const paragraphs = Array.from({ length: 200_000 }, (_, i) => `<w:p><w:r><w:t>${i}</w:t></w:r></w:p>`).join("");
    const read = readDocx(patch(base, "word/document.xml", (x) => x.replace("<w:body>", `<w:body><w:sdt><w:sdtContent>${paragraphs}</w:sdtContent></w:sdt>`)));
    expect(read.model.blocks.length).toBeGreaterThan(200_000);
    const row = [[para("x")]];
    expect(writeDocx({ blocks: [{ type: "table", rows: Array.from({ length: 150_000 }, () => row) }] }, { maxTotalBytes: 200_000_000 }).bytes.length).toBeGreaterThan(0);
  });
});

describe.skipIf(!hasLibreOffice && !libreOfficeRequired)("LibreOffice round trip", () => {
  const lo = hasLibreOffice ? new LibreOffice() : undefined;
  const records: RoundTripRecord[] = [];
  afterAll(() => {
    lo?.dispose();
    writeArtifact("docx-roundtrip", records);
  });

  test("8 LibreOffice opens our file, prints it, and its re-save reads back to the same model", { timeout: 300_000 }, () => {
    expect(hasLibreOffice).toBe(true);
    const { bytes } = writeDocx(fixture);
    const pdf = lo!.convert(bytes, "docx", "pdf");
    expect(strFromU8(pdf.subarray(0, 5))).toBe("%PDF-");
    const plain = strFromU8(lo!.convert(bytes, "docx", "txt:Text"));
    for (const word of ["Plan", "Details", "two-b", "deep", "Apples", "paragraphs", "ends.", "Claim", "holds."]) expect(plain).toContain(word);
    const resaved = lo!.convert(bytes, "docx", "docx");
    const back = readDocx(resaved);
    records.push({ format: "docx", case: "fixture", written: sha256(bytes), resaved: sha256(resaved), pdfBytes: pdf.length, equal: isDeepStrictEqual(back.model, normalizeTextDoc(fixture)) });
    expect(back.model).toEqual(normalizeTextDoc(fixture));
  });

  /** LibreOffice keeps no nested list of another kind when it saves DOCX: the sublist takes the kind of its parent. */
  const likeParent = (list: List, ordered = list.ordered): List => ({ ...list, ordered, items: list.items.map((i) => (i.sublist ? { ...i, sublist: likeParent(i.sublist, ordered) } : i)) });
  const forDocx = (doc: TextDoc): TextDoc => ({ ...doc, blocks: doc.blocks.map((b) => (b.type === "list" ? likeParent(b) : b)) });

  test.each([
    ["docx", "odt"],
    ["odt", "docx"],
  ] as const)("8 %s through LibreOffice to %s reads back to the same model", { timeout: 300_000 }, (from, to) => {
    const write = { docx: writeDocx, odt: writeOdt };
    const read = { docx: readDocx, odt: readOdt };
    const input = from === "odt" ? forDocx(fixture) : fixture;
    const { bytes } = write[from](input);
    const converted = lo!.convert(bytes, from, to);
    const back = read[to](converted);
    const expected = normalizeTextDoc(input);
    records.push({ format: `${from}->${to}`, case: "fixture", written: sha256(bytes), resaved: sha256(converted), pdfBytes: 0, equal: isDeepStrictEqual(back.model, expected) });
    expect(back.model).toEqual(expected);
  });
});
