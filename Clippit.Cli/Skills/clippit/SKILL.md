---
name: clippit
description: Work with OpenXml PowerPoint, Word, and Excel files using Clippit CLI: split/build/verify PPTX, build/compare/assemble DOCX, convert DOCX/XLSX to HTML, and create XLSX.
allowed-tools: Bash(clippit:*) Bash(dotnet:*)
---
<!-- clippit-skill-version: bundled -->

# Clippit CLI

Use `clippit` for deterministic OpenXml document operations. Prefer it over manual zip/XML edits for `.pptx`, `.docx`, and `.xlsx` files.

## Discover commands

```bash
clippit --help
clippit pptx --help
clippit word --help
clippit excel --help
```

Run command-specific help before using unfamiliar options:

```bash
clippit pptx split --help
clippit word simplify-markup --help
```

## Automation rules

- Validate generated Office files with the matching `verify` command.
- Use `--format json` when another step needs machine-readable output.
- Do not overwrite user documents unless explicitly requested; write a new output file.
- Prefer manifest scaffolding commands over hand-authoring manifests from scratch.
- Success payloads are stdout. Command errors are stderr JSON with a stable `code`.
- For full options, trust `clippit <command> --help`; this skill is a compact runbook, not the full docs.

## Validate files

```bash
clippit pptx verify deck.pptx --format json
clippit word verify document.docx --format json
clippit excel verify workbook.xlsx --format json
```

If validation fails, inspect stderr/stdout JSON and report the diagnostics instead of assuming the file is usable.

## PowerPoint workflows

Split a deck:

```bash
clippit pptx split deck.pptx --output slides --manifest
```

Extract selected slides:

```bash
clippit pptx split deck.pptx --slides 1,3,6-9 --output slides
```

Build a deck from a manifest:

```bash
clippit pptx build init --output deck-manifest.json
clippit pptx build run deck-manifest.json --output rebuilt.pptx
clippit pptx verify rebuilt.pptx --format json
```

## Word workflows

Build/merge documents:

```bash
clippit word build init --output word-build.json
clippit word build run word-build.json --output merged.docx
clippit word verify merged.docx --format json
```

Compare or consolidate revisions:

```bash
clippit word compare before.docx after.docx --output compared.docx
clippit word consolidate original.docx alice.docx bob.docx --output consolidated.docx
```

Assemble a template with XML data:

```bash
clippit word assemble template.docx data.xml --output assembled.docx
clippit word verify assembled.docx --format json
```

Clean markup safely:

```bash
clippit word accept-revisions draft.docx --output accepted.docx
clippit word simplify-markup document.docx --accept-revisions --remove-comments --output simplified.docx
```

Convert Word and HTML:

```bash
clippit word to-html document.docx --output document.html
clippit word from-html article.html --output article.docx
```

## Excel workflows

```bash
clippit excel verify workbook.xlsx --format json
clippit excel to-html workbook.xlsx --sheet "Sheet1" --output sheet.html
clippit excel create workbook.json --output workbook.xlsx
clippit excel verify workbook.xlsx --format json
```

## More details

Read these only when needed:

- `references/workflows.md` for multi-step PPTX, DOCX, and XLSX recipes.
- `references/manifests.md` for minimal manifest examples and schema guidance.
- `references/output.md` for JSON output, exit codes, and scripting patterns.
