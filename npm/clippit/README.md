# clippit

**Clippit CLI** is a command-line tool for working with OpenXml files (PowerPoint, Word, Excel).
This npm package provides native binaries for all platforms — no .NET runtime required.

## Installation

```bash
npm install -g clippit
```

## Quick Start

```bash
# Split a deck into individual slides
clippit pptx split presentation.pptx --output ./slides/

# Build a deck from a manifest
clippit pptx build run manifest.json --output result.pptx

# Validate a PPTX file
clippit pptx verify presentation.pptx

# Validate a DOCX file
clippit word verify document.docx

# Validate an XLSX file
clippit excel verify spreadsheet.xlsx

# Get JSON output for scripting
clippit pptx split presentation.pptx --format json
```

## Commands

| Command           | Description                                                                                                                             |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| `pptx split`      | Split a `.pptx` into individual single-slide files. Supports slide range selection (`--slides`) and manifest generation (`--manifest`). |
| `pptx build init` | Scaffold a deck manifest (JSON).                                                                                                        |
| `pptx build run`  | Assemble a `.pptx` from a deck manifest.                                                                                                |
| `pptx verify`     | Validate a PPTX — schema, relationships, markup compatibility, and sections.                                                            |
| `word verify`     | Validate a DOCX — schema and relationships.                                                                                             |
| `word to-html`    | Convert a DOCX to HTML/CSS.                                                                                                             |
| `word from-html`  | Convert HTML/CSS to a DOCX.                                                                                                             |
| `excel to-html`   | Convert an XLSX sheet, range, or table to HTML/CSS.                                                                                     |
| `excel verify`    | Validate an XLSX — schema and relationships.                                                                                            |
| `version`         | Print version information.                                                                                                              |

### Common flags

| Flag                  | Description                                                |
| --------------------- | ---------------------------------------------------------- |
| `--format json\|text` | Structured JSON or human-readable output (default: `text`) |
| `--quiet` / `-q`      | Suppress success output; exit codes still reflect result   |
| `--force`             | Overwrite existing output files                            |
| `-`                   | Use stdin / stdout for piped workflows                     |

## Machine-readable output

Success payloads → **stdout** (compact JSON when `--format json` or stdout is piped).
Command errors → **stderr** (compact JSON with a stable symbolic `code`).

Published JSON schemas for manifests and result payloads are available at
[`docs/schemas/`](https://github.com/sergey-tihon/Clippit/tree/main/docs/schemas).

## Supported platforms

| Platform    | Package                                  |
| ----------- | ---------------------------------------- |
| Windows x64 | `@sergey-tihon/clippit-bin-win32-x64`    |
| macOS x64   | `@sergey-tihon/clippit-bin-darwin-x64`   |
| macOS arm64 | `@sergey-tihon/clippit-bin-darwin-arm64` |
| Linux x64   | `@sergey-tihon/clippit-bin-linux-x64`    |

The correct binary package is installed automatically as an optional dependency.

## Full documentation

➡️ **[https://sergey-tihon.github.io/Clippit/cli.html](https://sergey-tihon.github.io/Clippit/cli.html)**

## dotnet tool

Prefer .NET? Install via: `dotnet tool install -g Clippit.Cli`

## License

MIT © [Sergey Tihon](https://github.com/sergey-tihon)
