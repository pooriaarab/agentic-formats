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
    expect(bad.stderr).toContain('unknown command "summary"');
  });

  test("the MCP server starts over stdio", async () => {
    const client = new Client({ name: "probe", version: "0.0.0" });
    await client.connect(new StdioClientTransport({ command: "node", args: [cli, "mcp"] }));
    try {
      expect(client.getServerVersion()).toMatchObject({ name: "agentic-formats", version });
    } finally {
      await client.close();
    }
  });

  test("the worker answers /health and refuses the scaffold routes", async () => {
    const health = await worker.fetch(new Request("https://x.test/health"), {});
    expect(await health.json()).toMatchObject({ ok: true, pack: "agentic-formats", version });
    expect((await worker.fetch(new Request("https://x.test/summarize", { method: "POST" }), {})).status).toBe(404);
  });
});
