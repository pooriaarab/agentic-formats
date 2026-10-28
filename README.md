# agentic-formats

[![npm version](https://img.shields.io/npm/v/agentic-formats)](https://www.npmjs.com/package/agentic-formats)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://opensource.org/license/mit)

[Package](https://www.npmjs.com/package/agentic-formats) · [Source](https://github.com/pooriaarab/agentic-formats) · [Support](https://github.com/pooriaarab/agentic-formats/issues)

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

> Status: early. This release holds the neutral models and the input guards. The file
> readers and writers land next, one format family per release. The package is not on npm yet.

## Install

```sh
npm install agentic-formats
```

## What works today

- **Neutral models.** `TextDoc` (headings, paragraphs with marks, lists, tables, images, links,
  footnotes), `Workbook` (sheets, cells with value, formula, and number format), and `Deck`
  (slides with text boxes, shapes, images, and notes). An app converts to these once and never
  touches zip or XML.
- **Guards for hostile files.** `readZip` and `parseXml` refuse a zip bomb, too many entries,
  an entry that is too large, an unsafe entry name, a duplicate entry name, a DOCTYPE, an
  undefined entity, XML that nests too deep, and a part that is too large. Every reader will use
  them. Limits are in `DEFAULT_LIMITS`, and you can lower or raise each one.
- **Normalizing.** `normalizeTextDoc` merges equal runs, so you can compare two documents.

```ts
import { readZip, parseXml, withLimits, FormatError } from "agentic-formats";

try {
  const files = readZip(bytes, withLimits({ maxTotalBytes: 20_000_000 }), "docx", (n) => n.endsWith(".xml"));
} catch (error) {
  if (error instanceof FormatError) console.error(error.code, error.message);
}
```

## Planned

- ODT and DOCX documents, ODS and XLSX spreadsheets, ODP and PPTX slide decks.
- A CLI (`agentic-formats convert in.odt out.docx`) and MCP tools.
- Round-trip tests against LibreOffice output.

## Limits

- No format has a reader or a writer yet.
- Lengths in the models are in points. Every reader reads UTF-8 XML only.
- A file is never partly read. A file that goes past a limit is rejected.

## Develop

```sh
npm ci
npm run ci:local
```

## License

MIT.
