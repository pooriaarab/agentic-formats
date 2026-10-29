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
 *  8 LibreOffice cannot open a file the CLI wrote
 */
import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readFileSync, rmSync, symlinkSync, truncateSync, writeFileSync, existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { isDeepStrictEqual } from "node:util";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { zipSync } from "fflate";
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

let dir: string;
const at = (name: string) => join(dir, name);
const run = (...args: string[]) => spawnSync("node", [cli, ...args], { encoding: "utf8", cwd: dir, env: { ...process.env, AGENTIC_FORMATS_ROOT: "" } });
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
      links: 2,
      images: 1,
      footnotes: 1,
      warnings: [],
    });
    expect(json(run("inspect", "sheet.xlsx").stdout)).toMatchObject({ kind: "sheet", sheets: [{ name: "Sales", cells: 4, formulas: 1 }, { name: "Two", cells: 1, formulas: 0 }] });
    expect(json(run("inspect", "deck.pptx").stdout)).toMatchObject({ kind: "deck", width: 720, height: 405, slides: 3, shapes: { textbox: 3, shape: 3, picture: 1 }, notes: 2 });
  });

  test("4 it refuses a broken file the same way", () => {
    const r = run("inspect", "junk.odt");
    expect(r.status).toBe(1);
    expect(r.stderr).toContain("INVALID_FILE");
  });
});

describe("mcp over stdio", () => {
  const connect = async (root?: string) => {
    const client = new Client({ name: "probe", version: "0.0.0" });
    await client.connect(new StdioClientTransport({ command: "node", args: [cli, "mcp"], cwd: dir, env: { ...(process.env as Record<string, string>), AGENTIC_FORMATS_ROOT: root ?? "" } }));
    return client;
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
    const client = await connect(root);
    try {
      expect(await call(client, "inspect_file", { path: "in.odt" })).toMatchObject({ isError: false });
      expect(await call(client, "convert_file", { input: "in.odt", output: "out.docx" })).toMatchObject({ isError: false });
      for (const path of ["../outside/secret.odt", join(outside, "secret.odt"), "escape/secret.odt", "/etc/passwd.odt"]) {
        const r = await call(client, "inspect_file", { path });
        expect(r.isError, path).toBe(true);
        expect(JSON.stringify(r.body), path).toMatch(/outside the allowed folder|file not found/);
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
