#!/usr/bin/env node
import { startMcp } from "./mcp.js";

startMcp({ root: process.env.AGENTIC_FORMATS_ROOT }).catch((error: unknown) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
