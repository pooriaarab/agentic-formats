---
name: agentic-formats
description: Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides. Keywords: odf,ooxml,odt,ods,odp,docx,xlsx,pptx,office,documents,import,export,agents,mcp,cli.
---

# agentic-formats

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

## Use the SDK

```js
import { convert, inspect, readOdt, writeDocx } from "agentic-formats";

const { model, warnings } = readOdt(bytes);
const { bytes: docx } = writeDocx(model);
const result = convert(bytes, "odt", "docx"); // { bytes, warnings, from, to }
```

## Use the CLI

```sh
agentic-formats convert report.odt report.docx
agentic-formats inspect report.docx
```

Output is JSON on stdout. A refusal exits with 1 and prints the reason on stderr.
`convert` replaces no file unless you pass `--force`.

## Use the MCP server

Start it with `agentic-formats mcp`. Call `inspect_file`, `read_model`, or `convert_file`.
Set `AGENTIC_FORMATS_ROOT` to limit every path to one folder.

## Rules

- Every file is hostile input. Keep the limits.
- Report a `FormatError` and every warning to the user.
- Replace a file only when the user asked.
