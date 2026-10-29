# agentic-formats agent guide

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

## Mental model

A file goes in. A neutral model comes out. A model goes in. A file comes out.
The models are `TextDoc`, `Workbook`, and `Deck`. Apps keep these models
and never write zip or XML code.

## The three faces

- SDK: `import { convert, inspect, readOdt, writeDocx } from "agentic-formats"`. Each reader returns
  `{ model, warnings }`. Each writer returns `{ bytes, warnings }`. `convert(bytes, from, to)` joins them.
  Documents: `readOdt`, `writeOdt`, `readDocx`, `writeDocx`. Spreadsheets: `readOds`, `writeOds`,
  `readXlsx`, `writeXlsx`. Slides: `readOdp`, `writeOdp`, `readPptx`, `writePptx`.
- CLI: `npx agentic-formats convert in.odt out.docx` and `npx agentic-formats inspect in.docx`.
- MCP: `npx agentic-formats mcp`, then call a tool below.

Keywords: odf,ooxml,odt,ods,odp,docx,xlsx,pptx,office,documents,import,export,agents,mcp,cli.

## MCP tools

| Tool | Use it for |
| --- | --- |
| inspect_file | Read a file and get counts and warnings. Proves the file is valid. |
| read_model | Read a file as neutral JSON. Image bytes show as their size. |
| convert_file | Convert to another format of the same kind. |

## The loop

1. Run `inspect_file` on the input. Stop if it returns an error.
2. Run `read_model` when you need the content.
3. Run `convert_file` with the output path. Pass `overwrite` only when the user asked to replace a file.
4. Read the warnings in the result. Each one names content that was lost. Tell the user.

## Rules

- Treat every file as hostile input. Use `readZip` and `parseXml` with their limits.
- Never fetch a remote resource that a file names.
- Report a `FormatError` as a failure. Do not retry with looser limits to get past it.
- Report each warning from a reader or writer. A warning names content that was lost.
- Do not replace a file unless the user asked. Never write over the input.
