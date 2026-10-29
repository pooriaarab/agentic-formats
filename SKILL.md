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

Output is JSON on stdout. A command-line error (an unknown command, or a wrong or missing flag or argument to `convert`, `inspect`, or `mcp`) exits with 2. Any other refusal exits with 1. Both print the reason on stderr.
`convert` replaces no file unless you pass `--force`.

## Use the MCP server

Start it with `agentic-formats mcp`. Call `inspect_file`, `read_model`, or `convert_file`.
The tools read and write only the working folder of the server, which is never `/` or the home folder (the server refuses to start there). Set `AGENTIC_FORMATS_ROOT` to
use another folder. Set `AGENTIC_FORMATS_UNSAFE_ANY_PATH=1` only when the user asked for any path.
A path outside the folder returns an "outside the allowed folder" error: do not look for a way around it.

## Rules

- Every file is hostile input. Keep the limits.
- Report a `FormatError` and every warning to the user.
- Replace a file only when the user asked.
