/*
 * Hostile input. Every reader in this package starts with readZip and parseXml, so these tests
 * drive both with files built to hurt. The ways a file can fail are written first; each has a case.
 *
 * Zip:  1 a bomb that packs past the ratio   2 one entry over the entry cap   3 entries over the
 *       total cap   4 too many entries   5 a name that leaves the package (`..`, root, drive,
 *       backslash, control character)   6 the same name twice, also in another case   7 a bad part
 *       that the reader does not want   8 bytes that are not a zip   9 an input over the size cap
 *       10 a writer that goes past the same caps
 * XML:  11 a DOCTYPE with entities (billion laughs)   12 an undefined entity (external reference)
 *       13 nesting past the cap   14 elements past the cap   15 a part past the size cap
 *       16 broken structure   17 a character XML forbids   18 an encoding other than UTF-8
 */
import { strToU8, zipSync } from "fflate";
import { describe, expect, test } from "vitest";
import { DEFAULT_LIMITS, FormatError, withLimits } from "./limits.js";
import { parseXml } from "./xml.js";
import { isSafeEntryName, readZip, writeZip } from "./zip.js";

const NS = { "urn:t": "t", "urn:d": "" };
const any = () => true;

/** A zip with the stored method and no checks, so a test can name entries as it likes. */
function rawZip(entries: [string, Uint8Array][]): Uint8Array {
  const parts: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  const u32 = (v: DataView, at: number, n: number) => v.setUint32(at, n, true);
  for (const [name, data] of entries) {
    const nameBytes = strToU8(name);
    const local = new Uint8Array(30 + nameBytes.length + data.length);
    const lv = new DataView(local.buffer);
    u32(lv, 0, 0x04034b50);
    lv.setUint16(4, 20, true);
    u32(lv, 18, data.length);
    u32(lv, 22, data.length);
    lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    local.set(data, 30 + nameBytes.length);
    const dir = new Uint8Array(46 + nameBytes.length);
    const dv = new DataView(dir.buffer);
    u32(dv, 0, 0x02014b50);
    dv.setUint16(4, 20, true);
    dv.setUint16(6, 20, true);
    u32(dv, 20, data.length);
    u32(dv, 24, data.length);
    dv.setUint16(28, nameBytes.length, true);
    u32(dv, 42, offset);
    dir.set(nameBytes, 46);
    parts.push(local);
    central.push(dir);
    offset += local.length;
  }
  const size = central.reduce((n, c) => n + c.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  u32(ev, 0, 0x06054b50);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  u32(ev, 12, size);
  u32(ev, 16, offset);
  const out = new Uint8Array(offset + size + 22);
  let at = 0;
  for (const p of [...parts, ...central, end]) (out.set(p, at), (at += p.length));
  return out;
}

const rejects = (run: () => unknown, code: string, message?: RegExp) => {
  try {
    run();
  } catch (error) {
    expect(error).toBeInstanceOf(FormatError);
    expect((error as FormatError).code).toBe(code);
    if (message) expect((error as FormatError).message).toMatch(message);
    return;
  }
  throw new Error("expected the call to throw");
};

const body = (s: string) => strToU8(s);

describe("zip guards", () => {
  test("1 a bomb that packs past the ratio is refused before it unpacks", () => {
    const bomb = zipSync({ "content.xml": new Uint8Array(40 * 1024 * 1024) }, { level: 9 });
    expect(bomb.length).toBeLessThan(200 * 1024);
    rejects(() => readZip(bomb, DEFAULT_LIMITS, "test", any), "LIMIT_EXCEEDED", /packs more than/);
  });

  test("2 one entry over the entry cap", () => {
    const file = rawZip([["a.xml", new Uint8Array(500)]]);
    rejects(() => readZip(file, withLimits({ maxEntryBytes: 499 }), "test", any), "LIMIT_EXCEEDED", /entry larger/);
    expect(Object.keys(readZip(file, withLimits({ maxEntryBytes: 500 }), "test", any))).toEqual(["a.xml"]);
  });

  test("3 entries that are each fine but pass the total cap", () => {
    const file = rawZip(["a", "b", "c"].map((n) => [`${n}.xml`, new Uint8Array(600)] as [string, Uint8Array]));
    rejects(() => readZip(file, withLimits({ maxTotalBytes: 1500 }), "test", any), "LIMIT_EXCEEDED", /unpacks to more/);
  });

  test("4 too many entries", () => {
    const file = rawZip(["a", "b", "c", "d"].map((n) => [`${n}.xml`, body("x")] as [string, Uint8Array]));
    rejects(() => readZip(file, withLimits({ maxEntries: 3 }), "test", any), "LIMIT_EXCEEDED", /more than 3 entries/);
  });

  test.each(["../evil.xml", "a/../../evil.xml", "/etc/passwd", "C:/x.xml", "a\\b.xml", "a/./b.xml", "a//b.xml", "nul\u0000.xml", ""])(
    "5 the name %j leaves the package",
    (name) => {
      expect(isSafeEntryName(name)).toBe(false);
      rejects(() => readZip(rawZip([["ok.xml", body("x")], [name, body("x")]]), DEFAULT_LIMITS, "test", any), "INVALID_FILE", /unsafe entry name/);
    },
  );

  test("5 a normal nested name is safe", () => {
    expect(isSafeEntryName("word/media/image1.png")).toBe(true);
    expect(isSafeEntryName("META-INF/")).toBe(true);
  });

  test("6 the same name twice, also in another case", () => {
    rejects(() => readZip(rawZip([["a.xml", body("1")], ["a.xml", body("2")]]), DEFAULT_LIMITS, "test", any), "INVALID_FILE", /twice/);
    rejects(() => readZip(rawZip([["a.xml", body("1")], ["A.XML", body("2")]]), DEFAULT_LIMITS, "test", any), "INVALID_FILE", /twice/);
  });

  test("7 a bad entry the reader does not want still fails the file", () => {
    const file = rawZip([["good.xml", body("x")], ["../skip.bin", body("x")]]);
    rejects(() => readZip(file, DEFAULT_LIMITS, "test", (n) => n === "good.xml"), "INVALID_FILE");
  });

  test("8 bytes that are not a zip", () => {
    rejects(() => readZip(body("<html>not a zip</html>"), DEFAULT_LIMITS, "test", any), "INVALID_FILE", /not a valid test/);
    rejects(() => readZip(new Uint8Array(), DEFAULT_LIMITS, "test", any), "INVALID_FILE");
  });

  test("9 an input over the size cap", () => {
    rejects(() => readZip(new Uint8Array(101), withLimits({ maxInputBytes: 100 }), "test", any), "LIMIT_EXCEEDED", /larger than 100/);
  });

  test("a clean zip unpacks only the wanted entries", () => {
    const file = writeZip([{ name: "mimetype", data: "x", store: true }, { name: "a/b.xml", data: "<r/>" }, { name: "skip.bin", data: "y" }], DEFAULT_LIMITS);
    const files = readZip(file, DEFAULT_LIMITS, "test", (n) => n !== "skip.bin");
    expect(Object.keys(files).sort()).toEqual(["a/b.xml", "mimetype"]);
    expect(new TextDecoder().decode(files["a/b.xml"])).toBe("<r/>");
  });

  test("10 a writer goes past the same caps, and the same model gives the same bytes", () => {
    const one = { name: "big.bin", data: new Uint8Array(200) };
    rejects(() => writeZip([one], withLimits({ maxEntryBytes: 100 })), "LIMIT_EXCEEDED", /larger than 100/);
    rejects(() => writeZip([one, { ...one, name: "b.bin" }], withLimits({ maxTotalBytes: 300 })), "LIMIT_EXCEEDED", /export is larger/);
    rejects(() => writeZip([{ name: "../x", data: "x" }], DEFAULT_LIMITS), "INVALID_MODEL");
    expect(writeZip([one], DEFAULT_LIMITS)).toEqual(writeZip([one], DEFAULT_LIMITS));
  });
});

describe("xml guards", () => {
  const parse = (xml: string, limits = DEFAULT_LIMITS) => parseXml(xml, limits, NS);

  test("11 a DOCTYPE with nested entities is refused, never expanded", () => {
    const laughs = `<?xml version="1.0"?><!DOCTYPE lolz [<!ENTITY lol "lol"><!ENTITY lol2 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">]><lolz>&lol2;</lolz>`;
    rejects(() => parse(laughs), "INVALID_FILE", /DOCTYPE/);
  });

  test("12 an undefined or external entity is refused", () => {
    rejects(() => parse(`<r>&xxe;</r>`), "INVALID_FILE", /not defined/);
    rejects(() => parse(`<r a="&xxe;"/>`), "INVALID_FILE", /not defined/);
  });

  test("13 nesting past the cap", () => {
    const deep = (n: number) => "<a>".repeat(n) + "</a>".repeat(n);
    expect(parse(deep(100)).name).toBe("a");
    rejects(() => parse(deep(101)), "LIMIT_EXCEEDED", /nests deeper/);
  });

  test("14 elements past the cap", () => {
    const flat = (n: number) => `<r>${"<c/>".repeat(n)}</r>`;
    expect(parse(flat(9), withLimits({ maxXmlElements: 10 })).children).toHaveLength(9);
    rejects(() => parse(flat(10), withLimits({ maxXmlElements: 10 })), "LIMIT_EXCEEDED", /more than 10 elements/);
  });

  test("15 a part past the size cap, as text and as bytes", () => {
    rejects(() => parse("<r>0123456789</r>", withLimits({ maxXmlBytes: 10 })), "LIMIT_EXCEEDED");
    rejects(() => parseXml(body("<r>0123456789</r>"), withLimits({ maxXmlBytes: 10 }), NS), "LIMIT_EXCEEDED");
  });

  test.each([
    ["two roots", "<a/><b/>"],
    ["an unclosed tag", "<a><b></a>"],
    ["a stray end tag", "</a>"],
    ["text outside the root", "hello<a/>"],
    ["an unquoted attribute", "<a b=c/>"],
    ["a < in an attribute", `<a b="<"/>`],
    ["an unclosed comment", "<a><!-- x</a>"],
    ["a CDATA outside the root", "<![CDATA[x]]><a/>"],
    ["empty input", ""],
    ["a tag that never ends", "<a b="],
  ])("16 broken structure: %s", (_name, xml) => {
    rejects(() => parse(xml), "INVALID_FILE");
  });

  test("17 a character reference XML forbids", () => {
    rejects(() => parse("<r>&#0;</r>"), "INVALID_FILE", /character reference/);
    rejects(() => parse("<r>&#xD800;</r>"), "INVALID_FILE");
  });

  test("18 an encoding other than UTF-8", () => {
    rejects(() => parse(`<?xml version="1.0" encoding="UTF-16"?><r/>`), "INVALID_FILE", /UTF-16/);
    expect(parse(`<?xml version="1.0" encoding="utf-8"?><r/>`).name).toBe("r");
  });

  test("a valid part reads: entities, CDATA, comments, and any prefix for a known namespace", () => {
    const root = parse(
      `<?xml version="1.0"?><!-- c --><x:root xmlns:x="urn:t" xmlns="urn:d" x:id="7" plain="&lt;&#65;&#x42;"><x:item>a &amp; b<![CDATA[ <raw> ]]></x:item><leaf/></x:root>`,
    );
    expect(root.name).toBe("t:root");
    expect(root.attrs).toEqual({ "t:id": "7", plain: "<AB" });
    expect(root.children).toEqual([{ name: "t:item", attrs: {}, children: ["a & b", " <raw> "] }, { name: "leaf", attrs: {}, children: [] }]);
  });
});
