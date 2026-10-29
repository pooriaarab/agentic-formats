# agentic-formats agent guide

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

## Mental model

A file goes in. A neutral model comes out. A model goes in. A file comes out.
The models are `TextDoc`, `Workbook`, and `Deck`. Apps keep these models
and never write zip or XML code.

## Faces

- SDK: `import { normalizeTextDoc, readZip } from "agentic-formats"`.
- CLI: `npx agentic-formats --help`.
- MCP: `npx agentic-formats mcp`.

Keywords: odf,ooxml,odt,ods,odp,docx,xlsx,pptx,office,documents,import,export,agents,mcp,cli.

## Rules

- Treat every file as hostile input. Use `readZip` and `parseXml` with their limits.
- Never fetch a remote resource that a file names.
- Report a `FormatError` as a failure. Do not retry with looser limits to get past it.
- Report each warning from a reader or writer. A warning names content that was lost.
