# Clippit manifest guidance

Prefer scaffolding manifests and editing the generated JSON rather than inventing the shape from memory.

```bash
clippit pptx build init --output deck.json
clippit word build init --output word-build.json
```

For authoritative contracts, use the published schemas in `docs/schemas/` or the repository documentation. This file only gives minimal examples.

## PPTX deck manifest

Typical flow:

```bash
clippit pptx build init --output deck.json
clippit pptx build run deck.json --output built.pptx
```

Minimal shape:

```json
{
  "$schema": "https://sergey-tihon.github.io/Clippit/schemas/deck-manifest.v1.json",
  "title": "Combined deck",
  "output": "combined.pptx",
  "deck": [
    "[Intro]",
    "slides/title.pptx",
    "slides/agenda.pptx",
    "[Appendix]",
    "appendix.pptx"
  ]
}
```

Section markers are strings in square brackets. Paths should be kept relative to the manifest when possible.

## Word build manifest

Typical flow:

```bash
clippit word build init --output word-build.json
clippit word build run word-build.json --output merged.docx
```

Minimal shape:

```json
{
  "$schema": "https://sergey-tihon.github.io/Clippit/schemas/word-build-manifest.v1.json",
  "output": "merged.docx",
  "entries": [
    "cover.docx",
    "[Chapter 1]",
    "chapter-1.docx",
    "[Chapter 2]",
    "chapter-2.docx"
  ]
}
```

## Excel workbook definition

Use with:

```bash
clippit excel create workbook.json --output workbook.xlsx
```

Minimal shape varies with workbook features. Check `docs/schemas/workbook-definition.v1.json` and prefer existing examples when available.

## Rules

- Do not embed full schemas in prompts or generated answers; point to schema files.
- Preserve `$schema` when editing manifests.
- Keep paths portable and relative unless the user asks for absolute paths.
- After generating Office output from a manifest, run the matching `verify` command.
