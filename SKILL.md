---
name: agentic-formats
description: Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides. Keywords: odf,ooxml,odt,ods,odp,docx,xlsx,pptx,office,documents,import,export,agents,mcp,cli.
---

# agentic-formats

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

## Use the SDK

```js
import { normalizeTextDoc } from "agentic-formats";

const doc = normalizeTextDoc({ blocks: [{ type: "paragraph", content: [{ type: "text", text: "Hello" }] }] });
```

## Use the CLI and MCP server

`agentic-formats --help` lists the commands. `agentic-formats mcp` starts the MCP server.
File commands and tools land in later releases.

## Rules

- Every file is hostile input. Keep the limits.
- Report a `FormatError` and every warning as failures.
