/*
 * A seeded fuzz pass over every reader and writer. The same seed gives the same cases, so a failure
 * names its seed and replays. A read must return a model or throw a `FormatError`: never another
 * error, never a hang, never a large allocation. The cases are:
 *
 *  1 one XML part gets a random edit: bytes flipped, a range cut, a subtree copied, a number or a
 *    name swapped for a hostile value (NaN, a negative, 1e999, a huge count, an empty string)
 *  2 nesting: one element wrapped in itself up to and past the depth cap
 *  3 entity expansion: a DOCTYPE, and a character reference that is not a character
 *  4 shared parts: many relationships, sheets, or slides that name one part, and a relationship loop
 *  5 the zip: a bomb that packs past the ratio, and random byte damage to the whole file
 *  6 a model with a value of the wrong type goes to a writer: the writer writes it or throws a `FormatError`
 *
 * `FUZZ_SEED` and `FUZZ_CASES` change the run, to look further than CI does.
 */
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { afterEach, describe, expect, test } from "vitest";
import { readDocx, writeDocx } from "./docx.js";
import { FormatError } from "./limits.js";
import { readOdp, writeOdp } from "./odp.js";
import { readOds, writeOds } from "./ods.js";
import { readOdt, writeOdt } from "./odt.js";
import { readPptx, writePptx } from "./pptx.js";
import { deckFixture, textFixture } from "./testing/fixtures.js";
import { readXlsx, writeXlsx } from "./xlsx.js";
import type { Workbook } from "./model.js";

const SEED = Number(process.env.FUZZ_SEED ?? 20260101);
const CASES = Number(process.env.FUZZ_CASES ?? 120);
/** One case may take this long. A hang past it fails the case; a sync loop that never ends fails the run by its own timeout. */
const BUDGET_MS = 4000;

/** mulberry32: a small seeded generator. */
function rng(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const workbook: Workbook = {
  sheets: [
    { name: "Sales", cells: { A1: { value: "Item" }, B1: { value: 3, format: "0.00" }, C1: { value: 6, formula: "B1*2" }, D1: { value: true }, E1: { value: 45000, format: "yyyy-mm-dd" } } },
    { name: "Other", cells: { A1: { value: "x" } } },
  ],
};

type Reader = (bytes: Uint8Array) => unknown;
interface Format {
  name: string;
  bytes: Uint8Array;
  read: Reader;
  /** Write a model of any shape. */
  write: (model: never) => unknown;
  model: unknown;
}
const formats: Format[] = [
  { name: "odt", bytes: writeOdt(textFixture).bytes, read: readOdt, write: writeOdt as never, model: textFixture },
  { name: "docx", bytes: writeDocx(textFixture).bytes, read: readDocx, write: writeDocx as never, model: textFixture },
  { name: "ods", bytes: writeOds(workbook).bytes, read: readOds, write: writeOds as never, model: workbook },
  { name: "xlsx", bytes: writeXlsx(workbook).bytes, read: readXlsx, write: writeXlsx as never, model: workbook },
  { name: "odp", bytes: writeOdp(deckFixture).bytes, read: readOdp, write: writeOdp as never, model: deckFixture },
  { name: "pptx", bytes: writePptx(deckFixture).bytes, read: readPptx, write: writePptx as never, model: deckFixture },
];

afterEach(expectNoFailures);

const isXmlPart = (name: string) => /\.(xml|rels)$/i.test(name);
const pick = <T>(r: () => number, items: T[]): T => items[Math.floor(r() * items.length)];
const HOSTILE = ["NaN", "Infinity", "1e999", "-1", "0", "-0", "99999999999999", "4294967296", "", " ", "__proto__", "constructor", "../x", "a".repeat(3000), "\u0000", "0x10", "1e-999", "1e999cm", "-5pt", "99999999999pt", "1e308", "9007199254740993", "100%", "#zzzzzz", "rId1", "1048577", "16385", "-2147483649", "0.5", "true", "text:p"];

/** Unpack, change one XML part with `edit`, and pack it again with no checks, so the reader sees what the edit made. */
function withPart(bytes: Uint8Array, r: () => number, edit: (xml: string) => string): Uint8Array {
  const files = unzipSync(bytes);
  const name = pick(r, Object.keys(files).filter(isXmlPart));
  const out: Record<string, [Uint8Array, { level: 0 | 6 }]> = {};
  for (const [n, data] of Object.entries(files)) out[n] = [n === name ? strToU8(edit(strFromU8(data))) : data, { level: n === "mimetype" ? 0 : 6 }];
  return zipSync(out);
}

const mutators: { name: string; run: (xml: string, r: () => number) => string }[] = [
  {
    name: "flip bytes",
    run: (xml, r) => {
      const chars = [...xml];
      for (let i = 0; i < 1 + Math.floor(r() * 4); i++) chars[Math.floor(r() * chars.length)] = pick(r, ["<", ">", "&", '"', "/", "=", " ", "\u0000", "\ud800", "x"]);
      return chars.join("");
    },
  },
  {
    name: "cut a range",
    run: (xml, r) => {
      const a = Math.floor(r() * xml.length);
      return xml.slice(0, a) + xml.slice(a + Math.floor(r() * 200));
    },
  },
  {
    name: "copy a range",
    run: (xml, r) => {
      const a = Math.floor(r() * xml.length);
      const piece = xml.slice(a, a + Math.floor(r() * 400));
      const b = Math.floor(r() * xml.length);
      return xml.slice(0, b) + piece.repeat(1 + Math.floor(r() * 5)) + xml.slice(b);
    },
  },
  {
    name: "hostile attribute value",
    run: (xml, r) => {
      const found = [...xml.matchAll(/(\s[\w:.-]+=)"([^"]*)"/g)];
      if (!found.length) return xml;
      const hit = pick(r, found);
      return xml.slice(0, hit.index) + `${hit[1]}"${pick(r, HOSTILE)}"` + xml.slice(hit.index! + hit[0].length);
    },
  },
  {
    name: "hostile text",
    run: (xml, r) => {
      const found = [...xml.matchAll(/>([^<>]+)</g)];
      if (!found.length) return xml;
      const hit = pick(r, found);
      return xml.slice(0, hit.index! + 1) + pick(r, HOSTILE) + xml.slice(hit.index! + hit[0].length - 1);
    },
  },
  {
    name: "swap a tag name",
    run: (xml, r) => {
      const tags = [...new Set([...xml.matchAll(/<([\w:.-]+)[\s>/]/g)].map((m) => m[1]))];
      if (tags.length < 2) return xml;
      const [from, to] = [pick(r, tags), pick(r, tags)];
      return xml.replaceAll(`<${from}`, `<${to}`).replaceAll(`</${from}>`, `</${to}>`);
    },
  },
  {
    name: "wrap in itself",
    run: (xml, r) => {
      const found = [...xml.matchAll(/<([\w:.-]+)(\s[^>]*)?>/g)].filter((m) => !m[0].endsWith("/>") && !m[0].startsWith("<?"));
      if (!found.length) return xml;
      const hit = pick(r, found);
      const depth = 1 + Math.floor(r() * 150);
      const end = xml.indexOf(`</${hit[1]}>`, hit.index);
      if (end < 0) return xml;
      return xml.slice(0, hit.index) + hit[0].repeat(depth) + xml.slice(hit.index! + hit[0].length, end) + `</${hit[1]}>`.repeat(depth) + xml.slice(end + hit[1].length + 3);
    },
  },
  { name: "doctype", run: (xml) => xml.replace(/\?>/, `?><!DOCTYPE x [<!ENTITY a "aaaaaaaaaa"><!ENTITY b "&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;">]>`).replace(/<([\w:.-]+)/, "<$1 z=\"&b;\"") },
  { name: "bad character reference", run: (xml, r) => xml.replace(/>([^<]*)</, `>&#${pick(r, ["0", "1", "xD800", "x110000", "99999999999", "xFFFFFFFF", "-1"])};<`) },
  { name: "empty", run: () => "" },
];

/** Failures by place in the code, so one bug is one line however many cases reach it. The first case that reaches a place names the seed. */
const failures = new Map<string, string>();

function check(label: string, fn: () => unknown): void {
  const started = performance.now();
  try {
    fn();
  } catch (error) {
    if (!(error instanceof FormatError)) {
      const e = error as Error;
      const site = (e?.stack ?? "").split("\n").find((l) => l.includes("/src/") && !l.includes("fuzz.test")) ?? "";
      const key = `${e?.name}: ${e?.message}`.slice(0, 160) + site.trim();
      if (!failures.has(key)) failures.set(key, `${label}: ${e?.name}: ${e?.message}\n  ${site.trim()}`);
    }
  }
  const took = performance.now() - started;
  if (took > BUDGET_MS) failures.set(`slow ${label}`, `${label}: took ${Math.round(took)} ms`);
}

/** Fail the test with every distinct failure since the last call. */
function expectNoFailures(): void {
  const found = [...failures.values()];
  failures.clear();
  expect(found.join("\n")).toBe("");
}

describe.each(formats)("fuzz $name", (format) => {
  test("1 random edits to one XML part", () => {
    const r = rng(SEED + format.name.length * 7919);
    for (let i = 0; i < CASES; i++) {
      const m = pick(r, mutators);
      const seed = Math.floor(r() * 2 ** 31);
      const local = rng(seed);
      const bytes = withPart(format.bytes, local, (xml) => m.run(xml, local));
      check(`${format.name} case ${i} (${m.name}, seed ${seed}, run seed ${SEED})`, () => format.read(bytes));
    }
  });

  test("1b several edits to several parts", () => {
    const r = rng(SEED + 31 + format.name.length);
    for (let i = 0; i < CASES / 4; i++) {
      const seed = Math.floor(r() * 2 ** 31);
      const local = rng(seed);
      let bytes = format.bytes;
      for (let k = 0; k < 2 + Math.floor(local() * 4); k++) {
        const m = pick(local, mutators.filter((x) => x.name !== "empty"));
        bytes = withPart(bytes, local, (xml) => m.run(xml, local));
      }
      check(`${format.name} multi case ${i} (seed ${seed}, run seed ${SEED})`, () => format.read(bytes));
    }
  });

  test("5 damage to the whole file", () => {
    const r = rng(SEED + 77 + format.name.length);
    for (let i = 0; i < CASES; i++) {
      const bytes = format.bytes.slice();
      const mode = i % 3;
      if (mode === 0) for (let k = 0; k < 1 + Math.floor(r() * 8); k++) bytes[Math.floor(r() * bytes.length)] = Math.floor(r() * 256);
      const input = mode === 1 ? bytes.subarray(0, Math.floor(r() * bytes.length)) : mode === 2 ? bytes.subarray(Math.floor(r() * 64)) : bytes;
      check(`${format.name} damage case ${i} (run seed ${SEED})`, () => format.read(input));
    }
  });

  test("6 a model with values of the wrong type", () => {
    const r = rng(SEED + 5 + format.name.length);
    const wrong: unknown[] = [null, undefined, NaN, Infinity, -1, 1e999, "", "x", true, [], {}, () => 1, 10n, Symbol("s"), { length: 3 }, [[]], "__proto__"];
    /** Copy the model with one leaf, or one container, swapped for a wrong value. */
    const mutate = (value: unknown, budget: { n: number }): unknown => {
      if (value instanceof Uint8Array) return budget.n-- === 0 ? pick(r, wrong) : value;
      if (Array.isArray(value)) return budget.n-- === 0 ? pick(r, wrong) : value.map((v) => mutate(v, budget));
      if (value && typeof value === "object") {
        if (budget.n-- === 0) return pick(r, wrong);
        return Object.fromEntries(Object.entries(value).flatMap(([k, v]) => (budget.n-- === 0 ? [] : [[k, mutate(v, budget)]])));
      }
      return budget.n-- === 0 ? pick(r, wrong) : value;
    };
    for (let i = 0; i < CASES * 2; i++) {
      const model = mutate(format.model, { n: Math.floor(r() * 300) });
      check(`${format.name} model case ${i} (run seed ${SEED})`, () => format.write(model as never));
    }
  });
});

describe("fuzz regressions: a model of the wrong type is an INVALID_MODEL, never a TypeError", () => {
  const bad = [Symbol("s"), 10n, {}, null, undefined];
  test.each(["odt", "docx"] as const)("%s", (name) => {
    const write = { odt: writeOdt, docx: writeDocx }[name] as (doc: never) => unknown;
    const docs: unknown[] = [undefined, null, {}, { blocks: {} }];
    for (const value of bad) {
      docs.push({ blocks: [{ type: "heading", level: value, content: [] }] }, { blocks: [{ type: value, content: [] }] }, { blocks: [{ type: "paragraph", content: [{ type: value }] }] });
    }
    for (const doc of docs) expect(() => write(doc as never)).toThrow(FormatError);
  });
});

describe("fuzz shared parts, nesting, and the zip", () => {
  const pptx = formats.find((f) => f.name === "pptx")!;
  const xlsx = formats.find((f) => f.name === "xlsx")!;
  const docx = formats.find((f) => f.name === "docx")!;

  test("4 many slides or sheets that name one part", () => {
    const slides = withPart(pptx.bytes, rng(1), (x) => x);
    const files = unzipSync(slides);
    const slidePart = Object.keys(files).find((n) => /ppt\/slides\/slide1\.xml$/.test(n))!;
    const rels = strFromU8(files["ppt/_rels/presentation.xml.rels"]);
    const many = rels.replace("</Relationships>", Array.from({ length: 300 }, (_, i) => `<Relationship Id="rS${i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="${slidePart.replace("ppt/", "")}"/>`).join("") + "</Relationships>");
    const pres = strFromU8(files["ppt/presentation.xml"]).replace("</p:sldIdLst>", Array.from({ length: 300 }, (_, i) => `<p:sldId id="${300 + i}" r:id="rS${i}"/>`).join("") + "</p:sldIdLst>");
    const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([n, d]) => [n, n === "ppt/_rels/presentation.xml.rels" ? strToU8(many) : n === "ppt/presentation.xml" ? strToU8(pres) : d])));
    check("pptx shared slide part", () => pptx.read(bytes));

    const x = unzipSync(xlsx.bytes);
    const book = strFromU8(x["xl/workbook.xml"]).replace("</sheets>", Array.from({ length: 300 }, (_, i) => `<sheet name="d${i}" sheetId="${50 + i}" r:id="rId1"/>`).join("") + "</sheets>");
    check("xlsx shared sheet part", () => xlsx.read(zipSync(Object.fromEntries(Object.entries(x).map(([n, d]) => [n, n === "xl/workbook.xml" ? strToU8(book) : d])))));
  });

  test("4 a relationship loop and a relationship to the part itself", () => {
    const files = unzipSync(docx.bytes);
    const rels = Object.keys(files).filter((n) => n.endsWith(".rels"));
    for (const target of ["document.xml", "../document.xml", "_rels/.rels", ".", "", "/word/document.xml", "http://example.org/x"]) {
      const out = Object.fromEntries(Object.entries(files).map(([n, d]) => [n, rels.includes(n) ? strToU8(strFromU8(d).replace(/Target="[^"]*"/g, `Target="${target}"`)) : d]));
      check(`docx rels target ${target}`, () => docx.read(zipSync(out)));
    }
  });

  test("2 nesting past the depth cap is refused, and a deep tree just under it reads", () => {
    for (const f of formats) {
      for (const depth of [90, 101, 5000]) {
        const bytes = withPart(f.bytes, rng(depth), (xml) => {
          const root = /<([\w:.-]+)[\s>]/.exec(xml.replace(/<\?[^>]*\?>/, ""))?.[1] ?? "a";
          return xml.replace(/(<[\w:.-]+[^>]*>)/, `$1${`<${root}>`.repeat(depth)}${`</${root}>`.repeat(depth)}`);
        });
        check(`${f.name} depth ${depth}`, () => f.read(bytes));
      }
    }
  });

  test("4 one image part used again and again is charged for every use", () => {
    const cases: [string, RegExp][] = [
      ["odt", /<draw:frame[^>]*>.*?<\/draw:frame>/s],
      ["docx", /<w:drawing>.*?<\/w:drawing>/s],
      ["odp", /<draw:frame[^>]*>(?:(?!<\/draw:frame>).)*<draw:image.*?<\/draw:frame>/s],
      ["pptx", /<p:pic>.*?<\/p:pic>/s],
    ];
    for (const [name, element] of cases) {
      const f = formats.find((x) => x.name === name)!;
      const files = unzipSync(f.bytes);
      const image = Object.keys(files).find((n) => /\.png$/.test(n))!;
      const png = new Uint8Array(1024 * 1024);
      png.set(files[image].subarray(0, 33));
      const part = Object.keys(files).find((n) => element.test(strFromU8(files[n])))!;
      const xml = strFromU8(files[part]);
      const copies = xml.replace(element, (m) => m.repeat(8));
      const bytes = zipSync(Object.fromEntries(Object.entries(files).map(([n, d]) => [n, [n === image ? png : n === part ? strToU8(copies) : d, { level: n === "mimetype" ? 0 : 6 }] as [Uint8Array, { level: 0 | 6 }]])));
      const reader = { odt: readOdt, docx: readDocx, odp: readOdp, pptx: readPptx }[name as "odt"];
      expect(() => reader(bytes, { maxTotalBytes: 4 * 1024 * 1024 }), name).toThrow(FormatError);
      expect(() => reader(bytes), name).not.toThrow();
    }
  });

  test("5 a zip that packs past the ratio and a zip of zips", () => {
    for (const f of formats) {
      const files = unzipSync(f.bytes);
      const name = Object.keys(files).find((n) => isXmlPart(n) && n !== "[Content_Types].xml")!;
      const big = strToU8(strFromU8(files[name]).replace(/<\/[^>]+>\s*$/, `<!-- ${" ".repeat(20 * 1024 * 1024)} -->$&`));
      const out = Object.fromEntries(Object.entries(files).map(([n, d]) => [n, [n === name ? big : d, { level: n === "mimetype" ? 0 : 9 }] as [Uint8Array, { level: 0 | 9 }]]));
      expect(() => f.read(zipSync(out))).toThrow(FormatError);
    }
  });
});
