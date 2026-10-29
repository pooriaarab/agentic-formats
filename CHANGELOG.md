# Changelog

## Unreleased

- The XLSX reader returns at most 50 warnings and one "N more warnings" line, and it names more dropped content: named ranges, hidden sheets, charts and comments, conditional formats, array formulas, rich text, values it cannot read, and missing styles.

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
