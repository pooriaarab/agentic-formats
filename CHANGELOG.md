# Changelog

## 0.4.2

- `agentic-formats mcp --help` and `-h` print the usage and exit 0. `--version` and `-V` print the version and exit 0. Before, the subcommand started the server, and the flag waited on stdin. An unknown option after `mcp` exits 2, as it does for `agentic-formats-mcp`.

## 0.4.1

- The README says a command-line error exits 2. That covers an unknown command, and an unknown flag or argument to `convert` or `inspect`. Every other error exits 1, such as a missing file, an unknown extension, or an output that exists. 0.4.0 said every error exits 1.
- The README names `ODT_MIME`, `ODS_MIME`, and `ODP_MIME`, the media types of the ODF formats. The code does not change.

## 0.4.0

- `agentic-formats-mcp --help` and `-h` print the usage and exit 0. `--version` and `-V` print the version and exit 0. An unknown argument exits 2. In 0.3.0 the MCP binary started the server for every argument, so `--help` waited on stdin. The `agentic-formats` CLI already behaved this way.
- The `zod` dependency is now 4.6.7. This fixes GHSA-mrhx-c7xc-m3w5, which 0.3.0 carried through `zod` 3.25. The MCP tools reject a bad argument with the same fields, but zod 4 words the message in a different way.
- Direct tests for the ODF and OOXML helpers in `odf.ts` and `ooxml.ts`. The behavior does not change.

## 0.3.0

- New `writePdf(model, limits?)` writes a `Deck` (one page per slide, at the deck size) or a `TextDoc` (Letter pages) as a PDF. It returns `{ bytes, warnings }`, like the other writers. It is plain JavaScript and runs in Node and in a Worker.
- Text, headings, nested lists, tables, footnotes, links, PNG (with alpha) and JPEG images, and the shapes of a slide are written. Text uses the standard Helvetica and Courier faces, so a character outside Latin-1 and the Windows punctuation is written as `?`.
- The writer reports text cut to fit a slide box, images it cannot embed (GIF, other bytes, interlaced PNG), characters it cannot hold, links that are not web or mail addresses, speaker notes, super and subscript, and a table inside a table, PNG transparency from a `tRNS` chunk, a list nested too deep for the page, and a table with more columns than fit, image descriptions, bold or italic on code text, and a title cut to 500 characters.
- The output is the same for the same model: a fixed creation date and fixed font names. It takes the `Limits` shape of the readers.
- `writePdf` is in the SDK only. The CLI and the MCP tools do not convert to PDF yet.

## 0.2.0

- A bad model fails with a `FormatError` (`INVALID_MODEL`), not a `TypeError`. This includes a block or inline type that is not a string, such as a symbol.
- `writeOdt` now checks the block list first. 0.1.0 threw a `TypeError` for a block list that was not a list.
- `writeOdt` and `writeDocx` refuse a heading level that is not a number. 0.1.0 wrote a numeric string, `null`, `undefined`, or an object as a heading level. These now fail.
- A seeded fuzz test runs every reader and writer.
- Output files are written to a temporary file in the same folder. Then a rename, or a hard link when `overwrite` is off, puts the file at the path. A failed write keeps the old file. A replaced file keeps its permission bits, not its owner or ACLs.
- The CLI refuses a symbolic link at the output path, as in 0.1.0. In folder mode, the MCP tools follow a link that stays inside the allowed folder.
- On a filesystem with no hard links, a new file is written in place, not atomically.
- A crash during a write can leave a hidden `.<name>.<hex>.tmp` file in the output folder. The name is cut to 32 characters.
- Three refusals changed. A read-only file (mode 0444) is now replaced when its folder can be written, as `mv` does. 0.1.0 refused it. An existing file in a folder that cannot be written is no longer written in place. The write fails. A file with a second name (a hard link) is now replaced, and the other name keeps the old bytes. 0.1.0 refused it.
- The XLSX reader adds lost-content warnings: named ranges, print areas, filters, hidden sheets, charts, images, comments, conditional formats, data validation, array formulas and data tables, rich text in a cell, cell values it cannot read, and missing styles. The reader lists at most 50 warnings that quote a name from the file, then one "N more warnings" line. A sheet name in a warning is cut to 60 characters.

## 0.1.0

First release.

- Neutral models: `TextDoc`, `Workbook`, and `Deck`.
- Input guards for hostile files: `readZip` and `parseXml` with `DEFAULT_LIMITS`.
- ODT: `readOdt` and `writeOdt`.
- DOCX: `readDocx` and `writeDocx`.
- ODS: `readOds` and `writeOds`.
- XLSX: `readXlsx` and `writeXlsx`.
- ODP: `readOdp` and `writeOdp`.
- PPTX: `readPptx` and `writePptx`.
- `convert` and `inspect` in the SDK, and the `agentic-formats` CLI.
- The `agentic-formats mcp` server with `inspect_file`, `read_model`, and `convert_file`.
- `normalizeTextDoc` to compare two documents.
