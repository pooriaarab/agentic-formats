# Changelog

## 0.2.0

- Writers refuse a bad model with a `FormatError` (`INVALID_MODEL`), not a `TypeError`. This applies to a block list that is not a list, a heading level that is not a number, and a block or inline item of an unknown type. `writeOdt` and `writeDocx` check these. A seeded fuzz test runs every reader and writer.
- Output files are written to a temporary file in the same folder, then renamed into place. A failed write keeps the old file. A replaced file keeps its mode, not its owner. A file with a second name (a hard link) is replaced; the other name keeps the old bytes. A symbolic link at the output path is refused.
- The XLSX reader adds lost-content warnings: named ranges, print areas, filters, hidden sheets, charts, images, comments, conditional formats, data validation, array formulas and data tables, rich text in a cell, cell values it cannot read, and missing styles. Warnings that quote a name from the file are capped at 50, followed by one "N more warnings" line.

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
