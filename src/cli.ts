#!/usr/bin/env node
import { realpathSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { VERSION } from "./version.js";

const HELP = `agentic-formats ${VERSION}
Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

Usage: agentic-formats <command>

Commands:
  mcp              Start the MCP server on stdio.

Flags:
  --help, -h       Print this help.
  --version, -V    Print the version.`;

export async function runCli(argv: string[]): Promise<number> {
  const [first] = argv;
  if (first === undefined || first === "--help" || first === "-h") {
    console.log(HELP);
    return 0;
  }
  if (first === "--version" || first === "-V") {
    console.log(VERSION);
    return 0;
  }
  if (first === "mcp") {
    const { startMcp } = await import("./mcp.js");
    await startMcp();
    return 0;
  }
  console.error(`unknown command "${first}"\nRun "agentic-formats --help" for usage.`);
  return 2;
}

function isMainModule(): boolean {
  const invoked = process.argv[1];
  if (typeof invoked !== "string") return false;
  try {
    return import.meta.url === pathToFileURL(realpathSync(invoked)).href;
  } catch {
    return false;
  }
}

if (isMainModule()) {
  runCli(process.argv.slice(2)).then(
    (code) => {
      // Exit nonzero at once. On zero, fall off the end so a running server keeps its stdio open.
      if (code !== 0) process.exit(code);
    },
    (error: unknown) => {
      console.error(error instanceof Error ? error.message : String(error));
      process.exit(1);
    },
  );
}
