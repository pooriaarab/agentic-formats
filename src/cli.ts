#!/usr/bin/env node
import { realpathSync } from "node:fs";
import { pathToFileURL } from "node:url";
import { FORMATS, convert, formatOfName, inspect, type Format } from "./convert.js";
import { UsageError, describeError, readInput, writeOutput } from "./files.js";
import { VERSION } from "./version.js";

const HELP = `agentic-formats ${VERSION}
Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

Usage: agentic-formats <command>

Commands:
  convert <in> <out> [--force]   Convert a file. The extensions pick the formats.
  inspect <file>                 Print what a file holds, as JSON.
  mcp                            Start the MCP server on stdio. Its tools read and write only
                                 the working folder; see AGENTIC_FORMATS_ROOT below. Its only
                                 options are --help and --version. Any other option exits 2.

Formats: ${FORMATS.join(", ")}. A file converts to a format of its own kind:
documents (odt, docx), spreadsheets (ods, xlsx), slide decks (odp, pptx).
Output is JSON on stdout. Errors go to stderr. A command-line error exits 2; any other error exits 1.

Environment (mcp only):
  AGENTIC_FORMATS_ROOT             Limit every path to this folder instead of the working folder.
  AGENTIC_FORMATS_UNSAFE_ANY_PATH  Set to 1 to allow any path. An empty variable is an error.

Flags:
  --force          Replace the output file if it exists.
  --help, -h       Print this help.
  --version, -V    Print the version.`;

const MCP_FLAGS = ["--help", "-h", "--version", "-V"];

function formatOf(path: string): Format {
  const format = formatOfName(path);
  if (!format) throw new UsageError(`"${path}" has no known extension (${FORMATS.join(", ")})`);
  return format;
}

export async function runCli(argv: string[]): Promise<number> {
  const [first, ...rest] = argv;
  if (first === undefined || first === "--help" || first === "-h") {
    console.log(HELP);
    return 0;
  }
  if (first === "--version" || first === "-V") {
    console.log(VERSION);
    return 0;
  }
  if (first === "mcp") {
    const unknown = rest.find((a) => !MCP_FLAGS.includes(a));
    if (unknown !== undefined) {
      console.error(`Usage error: unknown argument "${unknown}" for mcp.\nRun "agentic-formats --help" for usage.`);
      return 2;
    }
    if (rest.some((a) => a === "--help" || a === "-h")) {
      console.log(HELP);
      return 0;
    }
    if (rest.some((a) => a === "--version" || a === "-V")) {
      console.log(VERSION);
      return 0;
    }
    try {
      const { startMcpFromEnv } = await import("./mcp.js");
      await startMcpFromEnv();
    } catch (error) {
      const message = describeError(error);
      if (message === undefined) throw error;
      console.error(message);
      return 1;
    }
    return 0;
  }
  const force = rest.includes("--force");
  const paths = rest.filter((a) => a !== "--force");
  try {
    if (first === "convert" && paths.length === 2 && paths.every((p) => !p.startsWith("-"))) {
      const [input, output] = paths;
      const result = convert(readInput(input), formatOf(input), formatOf(output));
      writeOutput(output, result.bytes, force, input);
      console.log(JSON.stringify({ input, output, from: result.from, to: result.to, bytes: result.bytes.length, warnings: result.warnings }));
      return 0;
    }
    if (first === "inspect" && paths.length === 1 && !force && !paths[0].startsWith("-")) {
      console.log(JSON.stringify(inspect(readInput(paths[0]), formatOf(paths[0]))));
      return 0;
    }
  } catch (error) {
    const message = describeError(error);
    if (message === undefined) throw error;
    console.error(message);
    return 1;
  }
  console.error(`Usage error: "${argv.join(" ")}" is not a valid command.\nRun "agentic-formats --help" for usage.`);
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
