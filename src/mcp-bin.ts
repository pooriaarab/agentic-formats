#!/usr/bin/env node
import { describeError } from "./files.js";
import { startMcpFromEnv } from "./mcp.js";
import { VERSION } from "./version.js";

const HELP = `agentic-formats-mcp ${VERSION}
MCP server for agentic-formats. It serves the tools on stdio and starts with no arguments.

Usage: agentic-formats-mcp [--help | --version]

Environment:
  AGENTIC_FORMATS_ROOT             Limit every path to this folder instead of the working folder.
                                   The server refuses to start in / or in the home folder without it.
                                   An empty value is an error.
  AGENTIC_FORMATS_UNSAFE_ANY_PATH  Set to 1 to allow any path. Any other value is an error, and so
                                   is setting it together with AGENTIC_FORMATS_ROOT.

Flags:
  --help, -h       Print this help and exit.
  --version, -V    Print the version and exit.`;

const args = process.argv.slice(2);

if (args.includes("--help") || args.includes("-h")) {
  console.log(HELP);
} else if (args.includes("--version") || args.includes("-V")) {
  console.log(VERSION);
} else if (args.length > 0) {
  console.error(`agentic-formats-mcp: unknown argument "${args[0]}". Run with --help to see the options.`);
  process.exit(2);
} else {
  startMcpFromEnv().catch((error: unknown) => {
    console.error(describeError(error) ?? (error instanceof Error ? error.message : String(error)));
    process.exit(1);
  });
}
