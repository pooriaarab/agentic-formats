#!/usr/bin/env node
import { describeError } from "./files.js";
import { startMcpFromEnv } from "./mcp.js";

startMcpFromEnv().catch((error: unknown) => {
  console.error(describeError(error) ?? (error instanceof Error ? error.message : String(error)));
  process.exit(1);
});
