---
name: clippit
description: PPTX automation with Clippit CLI (split, build, verify). Also supports DOCX and XLSX workflows via references.
allowed-tools: Bash(clippit:*)
---
<!-- clippit-skill-version: bundled -->

# Clippit CLI

Use `clippit` for deterministic OpenXml operations. Prefer it over manual zip/XML edits for `.pptx`, `.docx`, and `.xlsx`.

## PowerPoint (primary)

Split a deck into single-slide files with a manifest:

```bash
clippit pptx split deck.pptx --output slides --manifest
```

Extract a subset of slides:

```bash
clippit pptx split deck.pptx --slides 1,3,6-9 --output subset
```

Build a deck from a manifest:

```bash
clippit pptx build init --output deck.json
clippit pptx build run deck.json --output built.pptx
clippit pptx verify built.pptx --format json
```

Verify any PPTX:

```bash
clippit pptx verify deck.pptx --format json
```

## Other formats

For DOCX (build, compare, assemble, HTML conversion) and XLSX (create, to-html), see `references/workflows.md`.

## Rules

- Run `clippit pptx <command> --help` for full options.
- Use `--format json` for machine-readable output.
- Prefer manifest scaffolding (`pptx build init`) over hand-authoring manifests.
- Validate generated files with the matching `verify` command.
- Success payloads go to stdout; errors are stderr JSON with a stable `code`.
- Do not overwrite user documents unless explicitly requested.

## References

- `references/workflows.md` — multi-step recipes for PPTX, DOCX, XLSX.
- `references/manifests.md` — manifest examples and schema guidance.
- `references/output.md` — JSON contracts, exit codes, scripting patterns.
