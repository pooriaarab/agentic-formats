# agentic-formats

[![npm version](https://img.shields.io/npm/v/agentic-formats)](https://www.npmjs.com/package/agentic-formats)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://opensource.org/license/mit)

[Package](https://www.npmjs.com/package/agentic-formats) · [Source](https://github.com/pooriaarab/agentic-formats) · [Support](https://github.com/pooriaarab/agentic-formats/issues)

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

> **Safe by default.** The MCP server reads and writes only the folder it starts in. It refuses to start in `/` or in your home folder: set `AGENTIC_FORMATS_ROOT` to the folder the tools may use. An agent
> that reads untrusted text must not reach your other files. To use another folder, set
> `AGENTIC_FORMATS_ROOT`. To allow any path, set `AGENTIC_FORMATS_UNSAFE_ANY_PATH=1`. An empty
> variable is an error.

> Status: 0.4.0. The neutral models, the input guards, all six formats (ODT, DOCX, ODS, XLSX, ODP,
> PPTX), the CLI, and the MCP server work. A replaced file is written atomically.
> See [CHANGELOG.md](CHANGELOG.md).

## Install

```sh
npm install agentic-formats
```

## Quick start

```sh
npx agentic-formats convert report.odt report.docx
npx agentic-formats inspect report.docx
npx agentic-formats mcp
```

The first line needs a `report.odt` in the current folder. Use any ODT file you have. To make one
from a Word file, run `npx agentic-formats convert draft.docx report.odt`. The second line reads
the `report.docx` that the first line writes. `mcp` starts the MCP server. It prints one line on
stderr and then waits for a client on stdin. To see the server's options, run
`npx -p agentic-formats agentic-formats-mcp --help`.

`convert` picks the formats from the file extensions and prints JSON: the files, the sizes, and
`warnings` (each one names content that was lost). It refuses to replace a file unless you pass
`--force`, and it never replaces its input. A file converts to a format of its own kind:
documents (`odt`, `docx`), spreadsheets (`ods`, `xlsx`), slide decks (`odp`, `pptx`).
Errors go to stderr. A command-line error exits with code 2: an unknown command, an unknown option
to `mcp`, or an unknown flag or argument to `convert` or `inspect`. Any other error exits with code 1, such as a missing file,
an unknown extension, or an output that exists.

The MCP server has three tools: `inspect_file`, `read_model` (the file as neutral JSON), and
`convert_file`. Every path stays inside one folder. With no variable set, that folder is the
working folder of the server. Set `AGENTIC_FORMATS_ROOT` to choose another folder (`/` allows the
whole disk). With no `AGENTIC_FORMATS_ROOT` and no `AGENTIC_FORMATS_UNSAFE_ANY_PATH`, the server
stops with an error at start if the working folder is `/` or your home folder. A path that leaves
the folder, by `..` or by a link, is refused. To allow any path, set
`AGENTIC_FORMATS_UNSAFE_ANY_PATH=1`; the server then says so on stderr. An empty variable, a value
other than `1`, or both variables together stops the server with an error. The CLI is not limited:
a person at a shell chooses the paths. Add the server to Claude Code with `.mcp.json`:

```json
{ "mcpServers": { "agentic-formats": { "command": "npx", "args": ["-y", "agentic-formats", "mcp"], "env": { "AGENTIC_FORMATS_ROOT": "/path/to/folder" } } } }
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
- **ODP and PPTX.** `readOdp`, `writeOdp`, `readPptx`, and `writePptx` handle slides with text
  boxes (paragraphs with marks and links), rectangles, ellipses, lines, pictures, and notes, on a
  slide of any size. A PPTX placeholder with no position of its own takes it from its layout, then
  its master. LibreOffice round trips cover both formats and the trip from one to the other.
- **MIME types.** `ODT_MIME`, `ODS_MIME`, and `ODP_MIME` hold the media type of each ODF
  format (`application/vnd.oasis.opendocument.text`, `...spreadsheet`, `...presentation`). Use them
  to set a `Content-Type` header. Each ODF reader also reads the `-template` variant, and
  `readOdt` also reads the `-web` variant.
- **PDF.** `writePdf(deck)` writes one page per slide, and `writePdf(textDoc)` writes Letter pages.
  It handles text with marks and links, headings, lists, tables, footnotes, PNG and JPEG images,
  and rectangles, ellipses, and lines. A slide text box takes the largest size from 18 to 8
  points that fits, and text that still does not fit is cut and reported. It has no PDF reader. The
  output is deterministic and the code is plain JavaScript, so it runs in a Worker. See Limits.
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

- An HTTP API as a Cloudflare Worker, and a small web UI. The Worker answers `/health` only.
- Writing a file from a model that an agent builds (`write_model`), and more formats (Markdown, CSV), and PDF in the CLI and MCP tools.

## Limits

- Number formats: ODS keeps `0`, `0.00`, `#,##0`, `#,##0.00`, `0%`, `0.00%`, and `yyyy-mm-dd`
  (and more decimals). XLSX keeps any format code as text. Another format in ODS is reported in
  `warnings` and the value stays.
- Spreadsheets: no merged cells, styles, charts, or pivot tables. A formula that uses a reference
  outside A1 notation (a whole column, an external file) is read as its value, with a warning.
  A shared XLSX formula is read as its value. Time cells read as a fraction of a day.
- ODT has no support for merged cells, comments, tracked changes, text boxes, or styles beyond the
  marks in the model. A reader lists each one it skips in `warnings`.
- Slides: no groups, charts, tables, animations, themes, or master slide content. A shape holds
  no text. A shape that holds text reads as a text box. A line runs from the top left of its frame
  to the bottom right. The writers add a blank layout and a plain master. PPTX is checked against
  LibreOffice, not PowerPoint.
- A link run carries no underline mark: a viewer underlines a link by itself. A footnote does not
  start with a space. Two tables in a row merge in Word.
- LibreOffice keeps no nested list of another kind when it saves DOCX. The sublist takes the kind
  of its parent. Our DOCX writer and reader keep the kind.
- A list item holds one paragraph and an optional sublist. Extra paragraphs join with line breaks.
- Images are PNG, JPEG, or GIF, found by their first bytes. A remote image is never loaded.
- The CLI and the MCP tools read and write files on disk. They are not an HTTP service.
- Lengths in the models are in points. Every reader reads UTF-8 XML only.
- PDF: the standard Helvetica and Courier faces, so a character outside Latin-1 and Windows
  punctuation is written as `?` and reported. No GIF images, speaker notes, or superscript, and a
  table row taller than a page is cut. Footnotes are written as notes at the end of the document, not on the page that cites them. A heading can end a page, and a table header row is not repeated on the next page. The SDK writes PDF; the CLI and MCP tools do not yet.
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
