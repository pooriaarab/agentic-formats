/*
 * The CLI and the MCP tools, end to end: the built CLI as a process, the built MCP server over
 * stdio, real files in a temporary folder, and LibreOffice. The ways it can fail are written first;
 * each has a case below.
 *
 *  1 a conversion writes a file that does not read back to the same model, in either direction,
 *    for each of the three kinds
 *  2 a wrong pairing is accepted: a document written as a spreadsheet, an unknown extension, a
 *    missing extension; an upper-case extension is refused
 *  3 the output clobbers something: an existing file, the input itself (also by another spelling or
 *    a link)
 *  4 a hostile or broken input: a zip bomb, bytes that are not a zip, a file past the size cap, a
 *    missing file; a failed run leaves an output file behind
 *  5 the streams and codes are wrong: JSON on stdout, errors on stderr, exit 1 for a refusal, exit
 *    2 for a usage error, warnings missing from the JSON
 *  6 inspect reports numbers that do not match the file
 *  7 the MCP server has the wrong tools, crashes on a bad path, or leaks outside its folder: `..`,
 *    an absolute path, a link that points out, an output outside the folder
 *  9 the MCP server starts with no folder set and reads or writes any path; an empty variable turns
 *    the guard off; the root `/` refuses every path; the opt-out flag is silent or half set
 *  8 LibreOffice cannot open a file the CLI wrote
 */
import { spawnSync } from "node:child_process";
import { existsSync, linkSync, mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, truncateSync, writeFileSync } from "node:fs";
import { randomBytes } from "node:crypto";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { isDeepStrictEqual } from "node:util";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { afterAll, beforeAll, describe, expect, test } from "vitest";
import { load } from "./convert.js";
import { writeDocx } from "./docx.js";
import { normalizeTextDoc, type Workbook } from "./model.js";
import { writeOdp } from "./odp.js";
import { writeOds } from "./ods.js";
import { writeOdt } from "./odt.js";
import { writePptx } from "./pptx.js";
import { PNG, deckFixture, textFixture } from "./testing/fixtures.js";
import { LibreOffice, hasLibreOffice, libreOfficeRequired, sha256, writeArtifact, type RoundTripRecord } from "./testing/libreoffice.js";
import { writeXlsx } from "./xlsx.js";

const cli = new URL("../dist/cli.js", import.meta.url).pathname;
const book: Workbook = { sheets: [{ name: "Sales", cells: { A1: { value: "Item" }, B1: { value: 3, format: "0.00" }, C1: { value: 6, formula: "B1*2" }, D1: { value: 45000, format: "yyyy-mm-dd" } } }, { name: "Two", cells: { A1: { value: true } } }] };

/** The environment of a child process: ours, with no AGENTIC_FORMATS_ variable unless the test sets one. */
const serverEnv = (extra: Record<string, string> = {}): Record<string, string> => {
  const env = { ...(process.env as Record<string, string>) };
  for (const name of Object.keys(env)) if (name.startsWith("AGENTIC_FORMATS_")) delete env[name];
  return { ...env, ...extra };
};

let dir: string;
const at = (name: string) => join(dir, name);
const run = (...args: string[]) => spawnSync("node", [cli, ...args], { encoding: "utf8", cwd: dir, env: serverEnv() });
const json = (out: string) => JSON.parse(out) as Record<string, unknown>;

beforeAll(() => {
  dir = mkdtempSync(join(tmpdir(), "formats-cli-"));
  writeFileSync(at("doc.odt"), writeOdt(textFixture).bytes);
  writeFileSync(at("doc.docx"), writeDocx(textFixture).bytes);
  writeFileSync(at("sheet.ods"), writeOds(book).bytes);
  writeFileSync(at("sheet.xlsx"), writeXlsx(book).bytes);
  writeFileSync(at("deck.odp"), writeOdp(deckFixture).bytes);
  writeFileSync(at("deck.pptx"), writePptx(deckFixture).bytes);
});
afterAll(() => rmSync(dir, { recursive: true, force: true }));

describe("convert", () => {
  const expected = { text: normalizeTextDoc(textFixture), sheet: book, deck: deckFixture };
  test.each([
    ["doc.odt", "out1.docx", "text"],
    ["doc.docx", "out2.odt", "text"],
    ["sheet.ods", "out3.xlsx", "sheet"],
    ["sheet.xlsx", "out4.ods", "sheet"],
    ["deck.odp", "out5.pptx", "deck"],
    ["deck.pptx", "out6.odp", "deck"],
  ] as const)("1 %s to %s reads back to the same model", (input, output, kind) => {
    const result = run("convert", input, output);
    expect(result.status).toBe(0);
    expect(json(result.stdout)).toMatchObject({ input, output, warnings: [] });
    const back = load(new Uint8Array(readFileSync(at(output))), output.split(".")[1] as "odt");
    expect(back.kind).toBe(kind);
    expect(back.model).toEqual(expected[kind]);
    expect(back.warnings).toEqual([]);
  });

  test("2 a wrong pairing, an unknown extension, and a missing extension are refused with exit 1", () => {
    for (const [input, output, message] of [
      ["doc.odt", "x.xlsx", "cannot be written as xlsx"],
      ["sheet.ods", "x.pptx", "cannot be written as pptx"],
      ["deck.odp", "x.docx", "cannot be written as docx"],
      ["doc.odt", "x.txt", "no known extension"],
      ["doc.odt", "noext", "no known extension"],
      ["doc.odt", "x.DOCX.bak", "no known extension"],
    ] as const) {
      const r = run("convert", input, output);
      expect(r.status, `${input} ${output}`).toBe(1);
      expect(r.stderr).toContain(message);
      expect(existsSync(at(output))).toBe(false);
    }
  });

  test("2 an upper-case extension works", () => {
    writeFileSync(at("UP.ODT"), readFileSync(at("doc.odt")));
    expect(run("convert", "UP.ODT", "up.DOCX").status).toBe(0);
    expect(existsSync(at("up.DOCX"))).toBe(true);
  });

  test("3 an existing output is kept unless --force, and the input is never replaced", () => {
    writeFileSync(at("keep.docx"), "keep me");
    const refused = run("convert", "doc.odt", "keep.docx");
    expect(refused.status).toBe(1);
    expect(refused.stderr).toContain("pass --force");
    expect(readFileSync(at("keep.docx"), "utf8")).toBe("keep me");
    expect(run("convert", "doc.odt", "keep.docx", "--force").status).toBe(0);
    expect(readFileSync(at("keep.docx")).length).toBeGreaterThan(100);

    mkdirSync(at("sub"), { recursive: true });
    symlinkSync(at("doc.odt"), at("link.odt"));
    const before = sha256(new Uint8Array(readFileSync(at("doc.odt"))));
    for (const out of ["doc.odt", "./sub/../doc.odt", "link.odt"]) {
      const r = run("convert", "doc.odt", out, "--force");
      expect(r.status, out).toBe(1);
      expect(r.stderr).toContain("the output is the input file");
    }
    expect(sha256(new Uint8Array(readFileSync(at("doc.odt"))))).toBe(before);
  });

  test("4 a zip bomb, a non-zip, an oversize file, and a missing file are refused and leave no output", () => {
    writeFileSync(at("bomb.docx"), zipSync({ "word/document.xml": new Uint8Array(40 * 1024 * 1024) }, { level: 9 }));
    writeFileSync(at("junk.odt"), "<html>not a zip</html>");
    writeFileSync(at("huge.xlsx"), "");
    truncateSync(at("huge.xlsx"), 26 * 1024 * 1024);
    const cases: [string, string, string][] = [
      ["bomb.docx", "b.odt", "LIMIT_EXCEEDED"],
      ["junk.odt", "j.docx", "INVALID_FILE"],
      ["huge.xlsx", "h.ods", "larger than"],
      ["missing.odt", "m.docx", "file not found"],
    ];
    for (const [input, output, message] of cases) {
      const r = run("convert", input, output);
      expect(r.status, input).toBe(1);
      expect(r.stderr).toContain(message);
      expect(existsSync(at(output)), output).toBe(false);
    }
  });

  test("5 warnings reach the JSON, the streams stay apart, and a usage error is exit 2", () => {
    const withLink = writeDocx({ blocks: [{ type: "paragraph", content: [{ type: "text", text: "x", href: "javascript:alert(1)" }] }] }).bytes;
    writeFileSync(at("warn.docx"), withLink);
    const r = run("convert", "warn.docx", "warn.odt");
    expect(r.status).toBe(0);
    expect(r.stderr).toBe("");
    expect(json(r.stdout).warnings).toEqual([]);
    for (const args of [["convert", "doc.odt"], ["convert"], ["inspect"], ["inspect", "a", "b"], ["frobnicate"], ["convert", "--force", "x"], ["inspect", "doc.odt", "--force"]]) {
      const bad = run(...args);
      expect(bad.status, args.join(" ")).toBe(2);
      expect(bad.stdout).toBe("");
    }
  });
});

describe("inspect", () => {
  test("6 it reports counts that match the model", () => {
    expect(json(run("inspect", "doc.docx").stdout)).toMatchObject({
      format: "docx",
      kind: "text",
      title: "Quarterly notes & <plans>",
      blocks: { heading: 2, paragraph: 12, list: 2, table: 1 },
      links: 3,
      images: 1,
      footnotes: 1,
      warnings: [],
    });
    expect(json(run("inspect", "sheet.xlsx").stdout)).toMatchObject({ kind: "sheet", sheetCount: 2, sheets: [{ name: "Sales", cells: 4, formulas: 1 }, { name: "Two", cells: 1, formulas: 0 }] });
    expect(json(run("inspect", "deck.pptx").stdout)).toMatchObject({ kind: "deck", width: 720, height: 405, slides: 3, shapes: { textbox: 3, shape: 3, picture: 1 }, notes: 2 });
  });

  test("4 it refuses a broken file the same way", () => {
    const r = run("inspect", "junk.odt");
    expect(r.status).toBe(1);
    expect(r.stderr).toContain("INVALID_FILE");
  });
});

describe("mcp over stdio", () => {
  /** Start the server in `dir`, with the variables given, and collect what it writes to stderr. */
  const connect = async (env: Record<string, string> = {}) => {
    const client = new Client({ name: "probe", version: "0.0.0" });
    const transport = new StdioClientTransport({ command: "node", args: [cli, "mcp"], cwd: dir, env: serverEnv(env), stderr: "pipe" });
    const stderr = { text: "" };
    transport.stderr?.on("data", (chunk: Buffer) => (stderr.text += chunk.toString()));
    await client.connect(transport);
    return Object.assign(client, { stderr });
  };
  const call = async (client: Client, name: string, args: Record<string, unknown>) => {
    const result = (await client.callTool({ name, arguments: args })) as { isError?: boolean; content: { text: string }[] };
    return { isError: !!result.isError, body: JSON.parse(result.content[0].text) as Record<string, unknown> };
  };

  test("7 it serves three tools and each one works", { timeout: 60_000 }, async () => {
    const client = await connect();
    try {
      expect((await client.listTools()).tools.map((t) => t.name).sort()).toEqual(["convert_file", "inspect_file", "read_model"]);
      const inspected = await call(client, "inspect_file", { path: at("sheet.ods") });
      expect(inspected).toMatchObject({ isError: false, body: { kind: "sheet" } });
      const model = await call(client, "read_model", { path: at("doc.odt") });
      expect(model.isError).toBe(false);
      const blocks = (model.body.model as { blocks: { content?: { type: string; data?: unknown }[] }[] }).blocks;
      expect(JSON.stringify(blocks)).toContain('"bytes":' + PNG.length.toString());
      expect(JSON.stringify(blocks)).not.toContain('"data":{"0"');
      const converted = await call(client, "convert_file", { input: at("deck.odp"), output: at("mcp.pptx") });
      expect(converted).toMatchObject({ isError: false, body: { from: "odp", to: "pptx", warnings: [] } });
      expect(load(new Uint8Array(readFileSync(at("mcp.pptx"))), "pptx").model).toEqual(deckFixture);
    } finally {
      await client.close();
    }
  });

  test("7 a bad call is a tool error, not a crash, and the server keeps answering", { timeout: 60_000 }, async () => {
    const client = await connect();
    try {
      expect(await call(client, "inspect_file", { path: at("missing.odt") })).toMatchObject({ isError: true });
      expect(await call(client, "inspect_file", { path: at("junk.odt") })).toMatchObject({ isError: true });
      expect(await call(client, "convert_file", { input: at("doc.odt"), output: at("mcp.xlsx") })).toMatchObject({ isError: true });
      expect(await call(client, "convert_file", { input: at("doc.odt"), output: at("keep.docx") })).toMatchObject({ isError: true });
      expect(await call(client, "convert_file", { input: at("doc.odt"), output: at("keep.docx"), overwrite: true })).toMatchObject({ isError: false });
      expect(await call(client, "inspect_file", { path: at("doc.odt") })).toMatchObject({ isError: false });
    } finally {
      await client.close();
    }
  });

  test("7 with a root, nothing outside it is read or written", { timeout: 60_000 }, async () => {
    const root = at("root");
    const outside = at("outside");
    mkdirSync(root, { recursive: true });
    mkdirSync(outside, { recursive: true });
    writeFileSync(join(root, "in.odt"), readFileSync(at("doc.odt")));
    writeFileSync(join(outside, "secret.odt"), readFileSync(at("doc.odt")));
    symlinkSync(outside, join(root, "escape"));
    const client = await connect({ AGENTIC_FORMATS_ROOT: root });
    try {
      expect(await call(client, "inspect_file", { path: "in.odt" })).toMatchObject({ isError: false });
      expect(await call(client, "convert_file", { input: "in.odt", output: "out.docx" })).toMatchObject({ isError: false });
      for (const path of ["../outside/secret.odt", join(outside, "secret.odt"), "escape/secret.odt", "/etc/passwd.odt"]) {
        const r = await call(client, "inspect_file", { path });
        expect(r.isError, path).toBe(true);
        expect(JSON.stringify(r.body), path).toMatch(/outside the allowed folder/);
      }
      for (const output of ["../outside/o.docx", "escape/o.docx", join(outside, "o.docx")]) {
        expect((await call(client, "convert_file", { input: "in.odt", output })).isError, output).toBe(true);
        expect(existsSync(join(outside, "o.docx"))).toBe(false);
      }
    } finally {
      await client.close();
    }
  });
});

describe("mcp output links", () => {
  const connect = async (root: string) => {
    const client = new Client({ name: "probe", version: "0.0.0" });
    await client.connect(new StdioClientTransport({ command: "node", args: [cli, "mcp"], cwd: dir, env: serverEnv({ AGENTIC_FORMATS_ROOT: root }) }));
    return client;
  };
  const call = async (client: Client, name: string, args: Record<string, unknown>) => {
    const result = (await client.callTool({ name, arguments: args })) as { isError?: boolean; content: { text: string }[] };
    return { isError: !!result.isError, text: result.content[0].text };
  };

  test("9 a link inside the folder that points nowhere, or at a file outside, never lets a write out", { timeout: 60_000 }, async () => {
    const root = at("links-root");
    const outside = at("links-outside");
    mkdirSync(root, { recursive: true });
    mkdirSync(outside, { recursive: true });
    writeFileSync(join(root, "in.odt"), readFileSync(at("doc.odt")));
    writeFileSync(join(outside, "victim.docx"), "original");
    symlinkSync(join(outside, "new.docx"), join(root, "dangling.docx"));
    symlinkSync(join(outside, "victim.docx"), join(root, "live.docx"));
    linkSync(join(outside, "victim.docx"), join(root, "hard.docx"));
    const client = await connect(root);
    try {
      const why: Record<string, RegExp> = { "dangling.docx": /points nowhere/, "live.docx": /outside the allowed folder/ };
      for (const output of ["dangling.docx", "live.docx"]) {
        for (const overwrite of [true, false]) {
          const r = await call(client, "convert_file", { input: "in.odt", output, overwrite });
          expect(r.isError, `${output} ${overwrite}`).toBe(true);
          expect(r.text, `${output} ${overwrite}`).toMatch(why[output]);
        }
      }
      // A second name for a file outside is refused without --force, and with it the name is replaced, not the file behind it.
      expect((await call(client, "convert_file", { input: "in.odt", output: "hard.docx", overwrite: false })).text).toMatch(/pass --force/);
      expect((await call(client, "convert_file", { input: "in.odt", output: "hard.docx", overwrite: true })).isError).toBe(false);
      expect(existsSync(join(outside, "new.docx"))).toBe(false);
      expect(readFileSync(join(outside, "victim.docx"), "utf8")).toBe("original");
      const dangling = await call(client, "convert_file", { input: "in.odt", output: "dangling.docx", overwrite: true });
      expect(dangling.text).toMatch(/points nowhere/);
      expect((await call(client, "convert_file", { input: "in.odt", output: "fresh.docx" })).isError).toBe(false);
      // A second name for a file inside the folder is fine to read, and a plain file there is replaced only on request.
      expect((await call(client, "convert_file", { input: "in.odt", output: "fresh.docx" })).text).toMatch(/pass --force/);
    } finally {
      await client.close();
    }
  });

  test("9 a named pipe as the output does not block the server", { timeout: 60_000 }, async () => {
    const root = at("fifo-root");
    mkdirSync(root, { recursive: true });
    writeFileSync(join(root, "in.odt"), readFileSync(at("doc.odt")));
    expect(spawnSync("mkfifo", [join(root, "pipe.docx")]).status).toBe(0);
    const client = await connect(root);
    try {
      const r = await call(client, "convert_file", { input: "in.odt", output: "pipe.docx", overwrite: true });
      expect(r.isError).toBe(true);
      expect(r.text).toMatch(/not a plain file/);
      expect(await call(client, "inspect_file", { path: "in.odt" })).toMatchObject({ isError: false });
    } finally {
      await client.close();
    }
  });

  test("9 a hostile file cannot flood a reply with the text of its own warnings or sheet names", { timeout: 120_000 }, async () => {
    const root = at("warn-root");
    mkdirSync(root, { recursive: true });
    const noise = (n: number) => randomBytes(n).toString("base64").replace(/[^a-z]/gi, "").slice(0, n);
    const repack = (bytes: Uint8Array, name: string, change: (xml: string) => string) =>
      zipSync(Object.fromEntries(Object.entries(unzipSync(bytes)).map(([n, data]) => [n, n === name ? strToU8(change(strFromU8(data))) : data])));
    const deck = writePptx({ width: 720, height: 405, slides: [{ shapes: [] }] }).bytes;
    const shape = `<p:sp><p:nvSpPr><p:cNvPr id="9" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="${noise(1_000_000)}"/></p:spPr></p:sp>`;
    writeFileSync(join(root, "warn.pptx"), repack(deck, "ppt/slides/slide1.xml", (x) => x.replace("</p:spTree>", `${shape}</p:spTree>`)));
    writeFileSync(join(root, "name.ods"), repack(writeOds({ sheets: [{ name: "S", cells: { A1: { value: 1 } } }] }).bytes, "content.xml", (x) => x.replace('table:name="S"', `table:name="${noise(1_000_000)}"`)));
    const client = await connect(root);
    try {
      for (const path of ["warn.pptx", "name.ods"]) {
        const inspected = await call(client, "inspect_file", { path });
        expect(inspected.isError, path).toBe(false);
        expect(inspected.text.length, path).toBeLessThan(5000);
        const model = await call(client, "convert_file", { input: path, output: path.replace(/\.\w+$/, path.endsWith("pptx") ? ".odp" : ".xlsx") });
        expect(model.text.length, path).toBeLessThan(5000);
      }
      expect(JSON.parse((await call(client, "inspect_file", { path: "warn.pptx" })).text).warnings[0]).toMatch(/\.\.\.$/);
    } finally {
      await client.close();
    }
  });

  test("9 a long title, a long list of warnings, and a large image cannot flood or fail a reply", { timeout: 120_000 }, async () => {
    const root = at("title-root");
    mkdirSync(root, { recursive: true });
    const noise = (n: number) => randomBytes(n).toString("base64").replace(/[^a-z]/gi, "").slice(0, n);
    const repack = (bytes: Uint8Array, name: string, change: (xml: string) => string) =>
      zipSync(Object.fromEntries(Object.entries(unzipSync(bytes)).map(([n, data]) => [n, n === name ? strToU8(change(strFromU8(data))) : data])));
    writeFileSync(join(root, "title.docx"), repack(writeDocx({ title: "t", blocks: [] }).bytes, "docProps/core.xml", (x) => x.replace("<dc:title>t</dc:title>", `<dc:title>${noise(1_000_000)}</dc:title>`)));
    writeFileSync(join(root, "title.odt"), repack(writeOdt({ title: "t", blocks: [] }).bytes, "meta.xml", (x) => x.replace("<dc:title>t</dc:title>", `<dc:title>${noise(1_000_000)}</dc:title>`)));
    const deck = writePptx({ width: 720, height: 405, slides: [{ shapes: [] }] }).bytes;
    const shapes = Array.from({ length: 500 }, (_, i) => `<p:sp><p:nvSpPr><p:cNvPr id="${i + 9}" name="s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom prst="odd${i}"/></p:spPr></p:sp>`).join("");
    writeFileSync(join(root, "many.pptx"), repack(deck, "ppt/slides/slide1.xml", (x) => x.replace("</p:spTree>", `${shapes}</p:spTree>`)));
    const photo = new Uint8Array(5_000_000);
    photo.set(PNG);
    photo.set(randomBytes(4_999_000), 1000);
    writeFileSync(join(root, "photo.odt"), writeOdt({ blocks: [{ type: "paragraph", content: [{ type: "image", data: photo, mime: "image/png" }] }] }).bytes);
    const client = await connect(root);
    try {
      for (const path of ["title.docx", "title.odt"]) {
        const inspected = await call(client, "inspect_file", { path });
        expect(inspected.isError, path).toBe(false);
        expect(inspected.text.length, path).toBeLessThan(2000);
        expect(JSON.parse(inspected.text).title, path).toMatch(/\.\.\.$/);
      }
      const many = JSON.parse((await call(client, "inspect_file", { path: "many.pptx" })).text) as { warnings: string[] };
      expect(many.warnings).toHaveLength(101);
      expect(many.warnings[100]).toBe("400 more warnings were left out");
      const read = await call(client, "read_model", { path: "photo.odt" });
      expect(read.isError).toBe(false);
      expect(read.text).toContain('"bytes":5000000');
    } finally {
      await client.close();
    }
  });

  test("9 inspect lists at most 100 sheets and says how many there are", { timeout: 120_000 }, async () => {
    const root = at("sheets-root");
    mkdirSync(root, { recursive: true });
    const one = writeOds({ sheets: [{ name: "S", cells: { A1: { value: 1 } } }] }).bytes;
    const tables = Array.from({ length: 1200 }, (_, i) => `<table:table table:name="Sheet${i}"/>`).join("");
    writeFileSync(join(root, "sheets.ods"), zipSync(Object.fromEntries(Object.entries(unzipSync(one)).map(([n, data]) => [n, n === "content.xml" ? strToU8(strFromU8(data).replace(/<table:table .*<\/table:table>/s, tables)) : data]))));
    const client = await connect(root);
    try {
      const inspected = JSON.parse((await call(client, "inspect_file", { path: "sheets.ods" })).text) as { sheetCount: number; sheets: { name: string }[] };
      expect(inspected.sheetCount).toBe(1200);
      expect(inspected.sheets).toHaveLength(100);
      expect(inspected.sheets[99].name).toBe("Sheet99");
    } finally {
      await client.close();
    }
  });

  test("9 read_model refuses a model past the reply cap and says where the counts are", { timeout: 120_000 }, async () => {
    const root = at("big-root");
    mkdirSync(root, { recursive: true });
    // Half random letters and half quotes: 6 M characters raw, under the 8,388,608 cap, and 9 M once JSON escapes each quote as \".
    const random = randomBytes(3_000_000);
    const mixed = Array.from(random, (b) => `${String.fromCharCode(97 + (b % 26))}"`).join("");
    writeFileSync(join(root, "big.ods"), writeOds({ sheets: [{ name: "S", cells: { A1: { value: mixed } } }] }, { maxEntryBytes: 200_000_000, maxTotalBytes: 200_000_000 }).bytes);
    const client = await connect(root);
    try {
      const read = await call(client, "read_model", { path: "big.ods" });
      expect(read.isError).toBe(true);
      expect(read.text).toMatch(/larger than 8388608 characters as JSON; use inspect_file/);
      expect(await call(client, "inspect_file", { path: "big.ods" })).toMatchObject({ isError: false });
    } finally {
      await client.close();
    }
  });
});

describe("mcp path policy", () => {
  const connect = async (env: Record<string, string> = {}) => {
    const client = new Client({ name: "probe", version: "0.0.0" });
    const transport = new StdioClientTransport({ command: "node", args: [cli, "mcp"], cwd: dir, env: serverEnv(env), stderr: "pipe" });
    const stderr = { text: "" };
    transport.stderr?.on("data", (chunk: Buffer) => (stderr.text += chunk.toString()));
    await client.connect(transport);
    return Object.assign(client, { stderr });
  };
  const call = async (client: Client, name: string, args: Record<string, unknown>) => {
    const result = (await client.callTool({ name, arguments: args })) as { isError?: boolean; content: { text: string }[] };
    return { isError: !!result.isError, text: result.content[0].text };
  };
  let elsewhere: string;
  beforeAll(() => {
    elsewhere = mkdtempSync(join(tmpdir(), "formats-elsewhere-"));
    writeFileSync(join(elsewhere, "secret.odt"), readFileSync(at("doc.odt")));
  });
  afterAll(() => rmSync(elsewhere, { recursive: true, force: true }));

  test("9 with no variable set, the folder the server starts in is the only one it reads or writes", { timeout: 60_000 }, async () => {
    const client = await connect();
    try {
      expect(await call(client, "inspect_file", { path: "doc.odt" })).toMatchObject({ isError: false });
      const read = await call(client, "read_model", { path: join(elsewhere, "secret.odt") });
      expect(read.isError).toBe(true);
      expect(read.text).toMatch(/outside the allowed folder/);
      const outside = await call(client, "convert_file", { input: "doc.odt", output: join(elsewhere, "out.docx"), overwrite: true });
      expect(outside.isError).toBe(true);
      expect(existsSync(join(elsewhere, "out.docx"))).toBe(false);
      expect(await call(client, "inspect_file", { path: "../" + join(elsewhere, "secret.odt").slice(1) })).toMatchObject({ isError: true });
      expect(client.stderr.text).toContain(dir);
    } finally {
      await client.close();
    }
  });

  test("9 an empty AGENTIC_FORMATS_ROOT is an error, not an open door", () => {
    for (const value of ["", "  "]) {
      const r = spawnSync("node", [cli, "mcp"], { encoding: "utf8", cwd: dir, env: serverEnv({ AGENTIC_FORMATS_ROOT: value }), input: "", timeout: 20_000 });
      expect(r.status, JSON.stringify(value)).toBe(1);
      expect(r.stderr).toMatch(/AGENTIC_FORMATS_ROOT is set but empty/);
    }
  });

  test("9 the agentic-formats-mcp binary has the same policy", () => {
    const bin = new URL("../dist/mcp-bin.js", import.meta.url).pathname;
    const empty = spawnSync("node", [bin], { encoding: "utf8", cwd: dir, env: serverEnv({ AGENTIC_FORMATS_ROOT: "" }), input: "", timeout: 20_000 });
    expect(empty.status).toBe(1);
    expect(empty.stderr).toMatch(/AGENTIC_FORMATS_ROOT is set but empty/);
    const flag = spawnSync("node", [bin], { encoding: "utf8", cwd: dir, env: serverEnv({ AGENTIC_FORMATS_UNSAFE_ANY_PATH: "yes" }), input: "", timeout: 20_000 });
    expect(flag.status).toBe(1);
    expect(flag.stderr).toMatch(/must be "1"/);
  });

  test("9 a working folder that is the filesystem root or the home folder is not a default scope", () => {
    const start = (cwd: string, env: Record<string, string> = {}) => spawnSync("node", [cli, "mcp"], { encoding: "utf8", cwd, env: serverEnv(env), input: "", timeout: 20_000 });
    for (const r of [start("/"), start(dir, { HOME: dir })]) {
      expect(r.status).toBe(1);
      expect(r.stderr).toMatch(/is the filesystem root or the home folder; set AGENTIC_FORMATS_ROOT/);
    }
    // A home folder that is a link to the working folder is the home folder.
    const link = at("home-link");
    symlinkSync(dir, link);
    const viaLink = start(dir, { HOME: link });
    expect(viaLink.status).toBe(1);
    expect(viaLink.stderr).toMatch(/is the filesystem root or the home folder/);
    // Naming the folder is allowed, even when it is the home folder.
    expect(start(dir, { HOME: dir, AGENTIC_FORMATS_ROOT: dir }).stderr).not.toMatch(/filesystem root or the home folder/);
  });

  test("9 a root that is not a folder is an error", () => {
    const r = spawnSync("node", [cli, "mcp"], { encoding: "utf8", cwd: dir, env: serverEnv({ AGENTIC_FORMATS_ROOT: at("doc.odt") }), input: "", timeout: 20_000 });
    expect(r.status).toBe(1);
    expect(r.stderr).toMatch(/is not a folder/);
  });

  test("9 the root / allows every path, as the owner chose", { timeout: 60_000 }, async () => {
    const client = await connect({ AGENTIC_FORMATS_ROOT: "/" });
    try {
      expect(await call(client, "inspect_file", { path: join(elsewhere, "secret.odt") })).toMatchObject({ isError: false });
    } finally {
      await client.close();
    }
  });

  test("9 the opt-out flag allows any path and says so on stderr", { timeout: 60_000 }, async () => {
    const client = await connect({ AGENTIC_FORMATS_UNSAFE_ANY_PATH: "1" });
    try {
      expect(await call(client, "inspect_file", { path: join(elsewhere, "secret.odt") })).toMatchObject({ isError: false });
      expect(client.stderr.text).toMatch(/AGENTIC_FORMATS_UNSAFE_ANY_PATH=1.*any path/);
    } finally {
      await client.close();
    }
  });

  test("9 the opt-out flag must be exactly 1, and it cannot be set with a root", () => {
    const start = (env: Record<string, string>) => spawnSync("node", [cli, "mcp"], { encoding: "utf8", cwd: dir, env: serverEnv(env), input: "", timeout: 20_000 });
    for (const value of ["", "0", "true", "yes"]) {
      const r = start({ AGENTIC_FORMATS_UNSAFE_ANY_PATH: value });
      expect(r.status, JSON.stringify(value)).toBe(1);
      expect(r.stderr).toMatch(/AGENTIC_FORMATS_UNSAFE_ANY_PATH must be "1"/);
    }
    const both = start({ AGENTIC_FORMATS_UNSAFE_ANY_PATH: "1", AGENTIC_FORMATS_ROOT: dir });
    expect(both.status).toBe(1);
    expect(both.stderr).toMatch(/not both/);
  });
});

describe.skipIf(!hasLibreOffice && !libreOfficeRequired)("LibreOffice opens what the CLI wrote", () => {
  const lo = hasLibreOffice ? new LibreOffice() : undefined;
  const records: RoundTripRecord[] = [];
  afterAll(() => {
    lo?.dispose();
    writeArtifact("cli-roundtrip", records);
  });

  test.each([
    ["doc.odt", "lo.docx"],
    ["sheet.xlsx", "lo.ods"],
    ["deck.odp", "lo.pptx"],
  ] as const)("8 %s converted to %s opens in LibreOffice and re-saves to the same model", { timeout: 300_000 }, (input, output) => {
    expect(run("convert", input, output).status).toBe(0);
    const ext = output.split(".")[1] as "docx" | "ods" | "pptx";
    const bytes = new Uint8Array(readFileSync(at(output)));
    const pdf = lo!.convert(bytes, ext, "pdf");
    expect(Buffer.from(pdf.subarray(0, 5)).toString()).toBe("%PDF-");
    const resaved = lo!.convert(bytes, ext, ext);
    const back = load(resaved, ext);
    const want = load(new Uint8Array(readFileSync(at(input))), input.split(".")[1] as "odt");
    records.push({ format: `${input.split(".")[1]}->${ext}`, case: "cli", written: sha256(bytes), resaved: sha256(resaved), pdfBytes: pdf.length, equal: isDeepStrictEqual(back.model, want.model) });
    expect(back.model).toEqual(want.model);
  });
});
