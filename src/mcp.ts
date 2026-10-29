import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { FORMATS, convert, formatOfName, inspect, load, modelJson, type Format } from "./convert.js";
import { UsageError, describeError, pathPolicyFromEnv, readInput, resolvePath, writeOutput, type PathPolicy } from "./files.js";
import { VERSION } from "./version.js";

const text = (value: unknown, isError = false) => ({ content: [{ type: "text" as const, text: typeof value === "string" ? value : JSON.stringify(value) }], ...(isError ? { isError: true as const } : {}) });

/** Run a tool body. A failure the caller can fix comes back as a tool error; anything else still throws. */
async function guarded(body: () => unknown) {
  try {
    return text(await body());
  } catch (error) {
    const message = describeError(error);
    if (message === undefined) throw error;
    return text({ error: message }, true);
  }
}

export function createMcpServer(policy: PathPolicy): McpServer {
  const server = new McpServer({ name: "agentic-formats", version: VERSION });
  const formatOf = (path: string): Format => {
    const format = formatOfName(path);
    if (!format) throw new UsageError(`"${path}" has no known extension (${FORMATS.join(", ")})`);
    return format;
  };
  const path = z.string().min(1).describe(`A file path with one of these extensions: ${FORMATS.join(", ")}.`);

  server.registerTool(
    "inspect_file",
    { description: "Read a file and report what it holds: counts of blocks, sheets, or slides, and warnings. Proves the file is valid.", inputSchema: { path }, annotations: { readOnlyHint: true, openWorldHint: false } },
    ({ path: p }) => guarded(() => {
      const file = resolvePath(p, policy);
      return inspect(readInput(file), formatOf(file));
    }),
  );

  server.registerTool(
    "read_model",
    { description: "Read a file as a neutral JSON model: a text document, a workbook, or a slide deck. Image bytes are replaced by their size.", inputSchema: { path }, annotations: { readOnlyHint: true, openWorldHint: false } },
    ({ path: p }) => guarded(() => {
      const file = resolvePath(p, policy);
      const loaded = load(readInput(file), formatOf(file));
      return `{"kind":"${loaded.kind}","warnings":${JSON.stringify(loaded.warnings)},"model":${modelJson(loaded.model)}}`;
    }),
  );

  server.registerTool(
    "convert_file",
    {
      description: "Convert a file to another format of the same kind (odt/docx, ods/xlsx, odp/pptx). The extensions pick the formats. Replaces the output only when overwrite is true. Report every warning: each names content that was lost.",
      inputSchema: { input: path, output: path, overwrite: z.boolean().optional().describe("Replace the output file if it exists.") },
      annotations: { readOnlyHint: false, destructiveHint: false, idempotentHint: false, openWorldHint: false },
    },
    ({ input, output, overwrite }) => guarded(() => {
      const from = resolvePath(input, policy);
      const to = resolvePath(output, policy);
      const result = convert(readInput(from), formatOf(from), formatOf(to));
      writeOutput(to, result.bytes, overwrite ?? false, from);
      return { input: from, output: to, from: result.from, to: result.to, bytes: result.bytes.length, warnings: result.warnings };
    }),
  );

  return server;
}

export async function startMcp(policy: PathPolicy): Promise<void> {
  console.error(
    "anyPath" in policy
      ? "agentic-formats mcp: AGENTIC_FORMATS_UNSAFE_ANY_PATH=1, so every tool reads and writes any path this process can"
      : `agentic-formats mcp: every tool is limited to ${policy.folder}`,
  );
  await createMcpServer(policy).connect(new StdioServerTransport());
}

/** Start the server with the policy that the environment names. A bad variable rejects with a `UsageError` before anything starts. */
export const startMcpFromEnv = async (env: NodeJS.ProcessEnv = process.env, cwd: string = process.cwd()): Promise<void> => startMcp(pathPolicyFromEnv(env, cwd));
