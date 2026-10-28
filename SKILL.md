---
name: agentic-formats
description: Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides. Keywords: odf,ooxml,odt,ods,odp,docx,xlsx,pptx,office,documents,import,export,agents,mcp,cli.
---

# agentic-formats

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

## Use the SDK

```js
import { summarizeText, dispatchText } from "agentic-formats";

summarizeText("some text");
await dispatchText("some text", { dryRun: true });
```

## Use the CLI

```sh
agentic-formats summary --text "some text"
agentic-formats dispatch --text "some text" --dry-run
```

## Use the MCP server

Start the server with `agentic-formats mcp`. Call `summarize_text`
for stats, `dispatch_text` for routes, `get_cost_report` for costs.

## Rules

- Preview each route with dry run first.
- Report guard denials and adapter errors as failures.
- Read the cost report at the end.
