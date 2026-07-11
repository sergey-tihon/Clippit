# clippit

**Clippit CLI** is a command-line tool for working with OpenXml files (PowerPoint, Word, Excel).
It supports PPTX split/build/verify workflows, manifest-driven DOCX build/compare/consolidate workflows, DOCX template assembly and markup cleanup, DOCX↔HTML conversion, and XLSX create/verify workflows.
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

# Scaffold a Word build manifest and merge it into a DOCX
clippit word build init --output word-build.json
clippit word build run word-build.json --output merged.docx

# Compare two DOCX files with tracked revisions
clippit word compare before.docx after.docx --output compared.docx

# Consolidate multiple DOCX revisions into one tracked-changes file
clippit word consolidate original.docx alice.docx bob.docx --output consolidated.docx

# Assemble a DOCX template with XML data
clippit word assemble template.docx data.xml --output assembled.docx

# Accept all tracked revisions in a DOCX file
clippit word accept-revisions draft.docx

# Remove non-content markup from a DOCX file
clippit word simplify-markup document.docx --accept-revisions --remove-comments

# Convert DOCX to HTML
clippit word to-html document.docx

# Convert HTML to DOCX
clippit word from-html article.html --css styles.css

# Validate an XLSX file
clippit excel verify spreadsheet.xlsx

# Create an XLSX workbook from a JSON definition
clippit excel create workbook.json --output report.xlsx

# Get JSON output for scripting
clippit pptx split presentation.pptx --format json
```

## Commands

| Command           | Description                                                                                                                             |
| ----------------- | --------------------------------------------------------------------------------------------------------------------------------------- |
| `pptx split`      | Split a `.pptx` into individual single-slide files. Supports slide range selection (`--slides`) and manifest generation (`--manifest`). |
| `pptx build init` | Scaffold a deck manifest (JSON).                                                                                                        |
| `pptx build run`  | Assemble a `.pptx` from a deck manifest.                                                                                                |
| `word build init` | Scaffold a Word build manifest (JSON).                                                                                                  |
| `word build run`  | Assemble a `.docx` from a Word build manifest.                                                                                          |
| `pptx verify`           | Validate a PPTX — schema, relationships, markup compatibility, and sections.                                                            |
| `word assemble`         | Assemble a DOCX template with XML data.                                                                                                  |
| `word compare`          | Compare two DOCX files and produce a tracked-revision DOCX.                                                                             |
| `word consolidate`      | Combine multiple DOCX revisions into one tracked-changes DOCX.                                                                          |
| `word accept-revisions` | Accept all tracked revisions in a DOCX file.                                                                                            |
| `word simplify-markup`  | Remove non-content markup from a DOCX file.                                                                                             |
| `word verify`           | Validate a DOCX — schema and relationships.                                                                                             |
| `word to-html`    | Convert a DOCX to HTML/CSS.                                                                                                             |
| `word from-html`  | Convert HTML/CSS to a DOCX.                                                                                                             |
| `excel to-html`   | Convert an XLSX sheet, range, or table to HTML/CSS.                                                                                     |
| `excel create`    | Generate an `.xlsx` workbook from a JSON workbook definition.                                                                           |
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

| Platform      | Package                                   |
| ------------- | ----------------------------------------- |
| Windows x64   | `@sergey-tihon/clippit-bin-win32-x64`     |
| Windows arm64 | `@sergey-tihon/clippit-bin-win32-arm64`   |
| macOS x64     | `@sergey-tihon/clippit-bin-darwin-x64`    |
| macOS arm64   | `@sergey-tihon/clippit-bin-darwin-arm64`  |
| Linux x64     | `@sergey-tihon/clippit-bin-linux-x64`     |
| Linux arm64   | `@sergey-tihon/clippit-bin-linux-arm64`   |

The correct binary package is installed automatically as an optional dependency.

## Full documentation

➡️ **[https://sergey-tihon.github.io/Clippit/cli.html](https://sergey-tihon.github.io/Clippit/cli.html)**

## dotnet tool

Prefer .NET? Install via: `dotnet tool install -g Clippit.Cli`

## License

MIT © [Sergey Tihon](https://github.com/sergey-tihon)
