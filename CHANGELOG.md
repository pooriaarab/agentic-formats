# Changelog

## Unreleased

- A seeded fuzz test runs every reader and writer. `writeOdt` and `writeDocx` now throw `FormatError` (`INVALID_MODEL`) for a model with a wrong-typed block list, heading level, or block or inline type, not a `TypeError`.
- Output files are written to a temporary file and renamed, so a failed write keeps the old file.
- The XLSX reader returns at most 50 warnings that quote a name from the file, and one "N more warnings" line, and it names more dropped content: named ranges, print areas, hidden sheets, charts and comments, conditional formats, array formulas, rich text, values it cannot read, and missing styles.

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
