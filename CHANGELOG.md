# Changelog

## 0.2.0

- Writers refuse a bad model with a `FormatError` (`INVALID_MODEL`), not a `TypeError`. `writeOdt` and `writeDocx` refuse a block list that is not a list. Both refuse a heading level that is not a number. 0.1.0 wrote a numeric string, `null`, `undefined`, or an object as a heading level, so these now fail. A seeded fuzz test runs every reader and writer.
- Output files are written to a temporary file in the same folder, then renamed into place. A failed write keeps the old file. A replaced file keeps its mode, not its owner. Two refusals changed. A read-only file (mode 0444) is now replaced when its folder can be written, as `mv` does. An existing file in a folder that cannot be written is no longer written in place, and the write fails. A file with a second name (a hard link) is replaced, and the other name keeps the old bytes. A symbolic link at the output path is refused. On a filesystem with no hard links, a new file is written in place, not atomically.
- The XLSX reader adds lost-content warnings: named ranges, print areas, filters, hidden sheets, charts, images, comments, conditional formats, data validation, array formulas and data tables, rich text in a cell, cell values it cannot read, and missing styles. The reader lists at most 50 warnings that quote a name from the file, then one "N more warnings" line.

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
