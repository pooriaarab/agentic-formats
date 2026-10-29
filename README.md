# agentic-formats

[![npm version](https://img.shields.io/npm/v/agentic-formats)](https://www.npmjs.com/package/agentic-formats)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://opensource.org/license/mit)

[Package](https://www.npmjs.com/package/agentic-formats) · [Source](https://github.com/pooriaarab/agentic-formats) · [Support](https://github.com/pooriaarab/agentic-formats/issues)

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

> Status: early. The neutral models, the input guards, ODT, DOCX, ODS, and XLSX work. The other
> formats land next, one family per release. The package is not on npm yet.

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
- **ODT.** `readOdt` and `writeOdt` handle headings, paragraphs with marks, links, nested lists,
  tables with a header row, images, footnotes, and the document title. A round-trip test has
  LibreOffice open, print, and re-save our file, and the re-saved file reads back to the same model.
  Each reader and writer returns warnings that name what the format could not hold.
- **DOCX.** `readDocx` and `writeDocx` handle the same content as ODT: headings, marks, links,
  nested lists, tables with a header row, images, footnotes, and the title. A round trip through
  LibreOffice covers DOCX to DOCX, and DOCX to ODT and back.
- **ODS and XLSX.** `readOds`, `writeOds`, `readXlsx`, and `writeXlsx` handle several sheets,
  strings, numbers, booleans, formulas as text with their cached result, and number formats.
  Formulas use A1 notation in the model (`SUM(A1:B2,Sheet2!C3)`). The ODS code converts them to
  OpenFormula and back. Dates are serial numbers with the `yyyy-mm-dd` format. LibreOffice
  round trips cover both formats and the trip from one to the other.
- **Normalizing.** `normalizeTextDoc` merges equal runs, so you can compare two documents.

```ts
import { readOdt, writeOdt, FormatError } from "agentic-formats";

try {
  const { model, warnings } = readOdt(bytes, { maxTotalBytes: 20_000_000 });
  const { bytes: out } = writeOdt(model);
} catch (error) {
  if (error instanceof FormatError) console.error(error.code, error.message);
}
```

## Planned

- ODP and PPTX slide decks.
- A CLI (`agentic-formats convert in.odt out.docx`) and MCP tools.
- Round-trip tests against LibreOffice output.

## Limits

- Number formats: ODS keeps `0`, `0.00`, `#,##0`, `#,##0.00`, `0%`, `0.00%`, and `yyyy-mm-dd`
  (and more decimals). XLSX keeps any format code as text. Another format in ODS is reported in
  `warnings` and the value stays.
- Spreadsheets: no merged cells, styles, charts, or pivot tables. A formula that uses a reference
  outside A1 notation (a whole column, an external file) is read as its value, with a warning.
  A shared XLSX formula is read as its value. Time cells read as a fraction of a day.
- ODT has no support for merged cells, comments, tracked changes, text boxes, or styles beyond the
  marks in the model. A reader lists each one it skips in `warnings`.
- A link run carries no underline mark: a viewer underlines a link by itself. A footnote does not
  start with a space. Two tables in a row merge in Word.
- LibreOffice keeps no nested list of another kind when it saves DOCX. The sublist takes the kind
  of its parent. Our DOCX writer and reader keep the kind.
- A list item holds one paragraph and an optional sublist. Extra paragraphs join with line breaks.
- Images are PNG, JPEG, or GIF, found by their first bytes. A remote image is never loaded.
- Lengths in the models are in points. Every reader reads UTF-8 XML only.
- A file is never partly read. A file that goes past a limit is rejected.

## Develop

```sh
npm ci
npm run ci:local
```

The round-trip tests need LibreOffice (`soffice` on the path). Set `REQUIRE_LIBREOFFICE=1` to fail
when it is missing. Without it, those tests skip. A run writes the records in `artifacts/`.

## License

MIT.
