# Clippit CLI — `dotnet tool`

**Clippit CLI** is a command-line tool for working with OpenXml files (PowerPoint, Word, Excel), built on [Clippit](https://github.com/sergey-tihon/Clippit) — the .NET OpenXml PowerTools library. It supports PPTX split/build/verify workflows, DOCX template assembly, and DOCX↔HTML conversion.

## Installation

```bash
dotnet tool install -g Clippit.Cli
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

# Compare two DOCX files with tracked revisions
clippit word compare before.docx after.docx --output compared.docx

# Assemble a DOCX template with XML data
clippit word assemble template.docx data.xml --output assembled.docx

# Accept all tracked revisions in a DOCX file
clippit word accept-revisions draft.docx

# Convert DOCX to HTML
clippit word to-html document.docx

# Convert HTML to DOCX
clippit word from-html article.html --css styles.css

# Validate an XLSX file
clippit excel verify spreadsheet.xlsx

# Get JSON output for scripting
clippit pptx split presentation.pptx --format json
```

## Commands

| Command | Description |
|---------|-------------|
| `pptx split` | Split a `.pptx` into individual single-slide files. Supports slide range selection (`--slides`) and manifest generation (`--manifest`). |
| `pptx build init` | Scaffold a deck manifest (JSON). |
| `pptx build run` | Assemble a `.pptx` from a deck manifest. |
| `pptx verify` | Validate a PPTX — schema, relationships, markup compatibility, and sections. |
| `word assemble` | Assemble a DOCX template with XML data. |
| `word compare` | Compare two DOCX files and produce a tracked-revision DOCX. |
| `word accept-revisions` | Accept all tracked revisions in a DOCX file. |
| `word verify` | Validate a DOCX — schema and relationships. |
| `word to-html` | Convert a DOCX to HTML/CSS. |
| `word from-html` | Convert HTML/CSS to a DOCX. |
| `excel to-html` | Convert an XLSX sheet, range, or table to HTML/CSS. |
| `excel verify` | Validate an XLSX — schema and relationships. |
| `version` | Print version information. |

### Common flags

| Flag | Description |
|------|-------------|
| `--format json|text` | Structured JSON or human-readable output (default: `text`) |
| `--quiet` / `-q` | Suppress success output; exit codes still reflect result |
| `--force` | Overwrite existing output files |
| `-` | Use stdin / stdout for piped workflows |

## Machine-readable output

Success payloads → **stdout** (compact JSON when `--format json` or stdout is piped).  
Command errors → **stderr** (compact JSON with a stable symbolic `code`).

Published JSON schemas for manifests and result payloads are available at  
[`docs/schemas/`](https://github.com/sergey-tihon/Clippit/tree/main/docs/schemas).

## Full documentation

➡️ **[https://sergey-tihon.github.io/Clippit/cli.html](https://sergey-tihon.github.io/Clippit/cli.html)**

## License

MIT © [Sergey Tihon](https://github.com/sergey-tihon)
