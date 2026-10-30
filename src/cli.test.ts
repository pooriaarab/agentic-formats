import { spawnSync } from "node:child_process";
import { readFileSync } from "node:fs";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { StdioClientTransport } from "@modelcontextprotocol/sdk/client/stdio.js";
import { describe, expect, test } from "vitest";
import worker from "./worker.js";

const cli = new URL("../dist/cli.js", import.meta.url).pathname;
const version = (JSON.parse(readFileSync(new URL("../package.json", import.meta.url), "utf8")) as { version: string }).version;

describe("built package", () => {
  test("the CLI prints its version and help, and rejects an unknown command", () => {
    expect(spawnSync("node", [cli, "--version"], { encoding: "utf8" }).stdout.trim()).toBe(version);
    expect(spawnSync("node", [cli, "--help"], { encoding: "utf8" }).stdout).toContain("agentic-formats");
    const bad = spawnSync("node", [cli, "summary"], { encoding: "utf8" });
    expect(bad.status).toBe(2);
    expect(bad.stderr).toContain('"summary" is not a valid command');
  });

  test("the CLI mcp subcommand prints its help and version and exits 0 without starting the server", () => {
    for (const flag of ["--help", "-h"]) {
      const help = spawnSync("node", [cli, "mcp", flag], { encoding: "utf8", timeout: 5000 });
      expect(help.error).toBeUndefined();
      expect(help.status).toBe(0);
      expect(help.stdout).toContain("Usage: agentic-formats <command>");
    }
    for (const flag of ["--version", "-V"]) {
      const ver = spawnSync("node", [cli, "mcp", flag], { encoding: "utf8", timeout: 5000 });
      expect(ver.error).toBeUndefined();
      expect(ver.status).toBe(0);
      expect(ver.stdout.trim()).toBe(version);
    }
  });

  test("the CLI mcp subcommand refuses an unknown option with exit 2 and starts no server", () => {
    const bad = spawnSync("node", [cli, "mcp", "--bogus", "--help"], { encoding: "utf8", timeout: 5000 });
    expect(bad.error).toBeUndefined();
    expect(bad.status).toBe(2);
    expect(bad.stdout).toBe("");
    expect(bad.stderr).toContain('unknown argument "--bogus" for mcp');
  });

  test("the MCP bin prints its help and version and exits 0 without starting the server", () => {
    const bin = new URL("../dist/mcp-bin.js", import.meta.url).pathname;
    const help = spawnSync("node", [bin, "--help"], { encoding: "utf8", timeout: 5000 });
    expect(help.error).toBeUndefined();
    expect(help.status).toBe(0);
    expect(help.stdout).toContain("Usage: agentic-formats-mcp");
    const ver = spawnSync("node", [bin, "--version"], { encoding: "utf8", timeout: 5000 });
    expect(ver.error).toBeUndefined();
    expect(ver.status).toBe(0);
    expect(ver.stdout.trim()).toBe(version);
    for (const flag of ["-h", "-V"]) {
      const short = spawnSync("node", [bin, flag], { encoding: "utf8", timeout: 5000 });
      expect(short.error).toBeUndefined();
      expect(short.status).toBe(0);
      expect(short.stdout).toContain(flag === "-h" ? "Usage: agentic-formats-mcp" : version);
    }
  });

  test("the MCP bin accepts one leading mcp before help and version", () => {
    const bin = new URL("../dist/mcp-bin.js", import.meta.url).pathname;
    for (const flag of ["--help", "-h", "--version", "-V"]) {
      const direct = spawnSync("node", [bin, flag], { encoding: "utf8", timeout: 5000 });
      const prefixed = spawnSync("node", [bin, "mcp", flag], { encoding: "utf8", timeout: 5000 });
      for (const result of [direct, prefixed]) {
        expect(result.error).toBeUndefined();
        expect(result.status).toBe(0);
        expect(result.stderr).toBe("");
      }
      expect(prefixed.stdout).toBe(direct.stdout);
    }
    for (const args of [["mcp", "mcp", "--help"], ["--help", "mcp"]]) {
      const bad = spawnSync("node", [bin, ...args], { encoding: "utf8", timeout: 5000 });
      expect(bad.status).toBe(2);
      expect(bad.stdout).toBe("");
      expect(bad.stderr).toContain('unknown argument "mcp"');
    }
  });

  test("the MCP bin reports a bad environment in one line, not a stack trace, and refuses an unknown flag", () => {
    const bin = new URL("../dist/mcp-bin.js", import.meta.url).pathname;
    const bad = spawnSync("node", [bin], { encoding: "utf8", timeout: 5000, env: { ...process.env, AGENTIC_FORMATS_ROOT: "" } });
    expect(bad.error).toBeUndefined();
    expect(bad.status).toBe(1);
    expect(bad.stderr).toContain("AGENTIC_FORMATS_ROOT is set but empty");
    expect(bad.stderr).not.toContain("    at ");
    const flag = spawnSync("node", [bin, "--hlep"], { encoding: "utf8", timeout: 5000 });
    expect(flag.error).toBeUndefined();
    expect(flag.status).toBe(2);
    expect(flag.stderr).toContain('unknown argument "--hlep"');
    const mixed = spawnSync("node", [bin, "--bogus", "--help"], { encoding: "utf8", timeout: 5000 });
    expect(mixed.status).toBe(2);
    expect(mixed.stdout).toBe("");
  });

  test("the MCP server starts over stdio", async () => {
    const bin = new URL("../dist/mcp-bin.js", import.meta.url).pathname;
    for (const args of [[cli, "mcp"], [bin, "mcp"]]) {
      const client = new Client({ name: "probe", version: "0.0.0" });
      await client.connect(new StdioClientTransport({ command: "node", args }));
      try {
        expect(client.getServerVersion()).toMatchObject({ name: "agentic-formats", version });
      } finally {
        await client.close();
      }
    }
  });

  test("the worker answers /health and refuses the scaffold routes", async () => {
    const health = await worker.fetch(new Request("https://x.test/health"), {});
    expect(await health.json()).toMatchObject({ ok: true, pack: "agentic-formats", version });
    expect((await worker.fetch(new Request("https://x.test/summarize", { method: "POST" }), {})).status).toBe(404);
  });
});
