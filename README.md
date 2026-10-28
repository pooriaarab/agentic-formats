# agentic-formats

[![npm version](https://img.shields.io/npm/v/agentic-formats)](https://www.npmjs.com/package/agentic-formats)
[![MIT license](https://img.shields.io/badge/license-MIT-blue)](https://opensource.org/license/mit)

[Package](https://www.npmjs.com/package/agentic-formats) · [Source](https://github.com/pooriaarab/agentic-formats) · [Support](https://github.com/pooriaarab/agentic-formats/issues)

Read and write open office files: ODF and OOXML for documents, spreadsheets, and slides.

> Status: scaffold. This repository holds the starter code only. It is not on npm
> yet, and none of the features below work yet.

## Planned features

- Read and write ODT and DOCX documents.
- Read and write ODS and XLSX spreadsheets.
- Read and write ODP and PPTX slide decks.
- One neutral document model per kind, so apps convert once.
- Treat every file as hostile input: size caps, zip-bomb checks, no external entities.
- Round-trip tests against LibreOffice output.

## Surfaces

One package gives you each way in:

- **Library and SDK**: `import { ... } from "agentic-formats"`.
- **CLI**: `npx agentic-formats --help`.
- **MCP server**: `npx agentic-formats mcp`, for Claude Code and other MCP clients.
- **HTTP API**: a Cloudflare Worker you deploy to your own account.
- **Web UI**: a small app served by the same Worker.

## Develop

```sh
npm ci
npm run build && npm run typecheck && npm test
```

## License

MIT.
