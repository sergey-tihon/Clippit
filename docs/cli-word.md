# Word CLI (`word`)

Use this page for Word (`.docx`) command reference.

- Back to [CLI start page](cli.md)
- Shared verify diagnostics schema: [schemas/README.md](schemas/README.md)

## Commands at a glance

| Command | Purpose |
| --- | --- |
| `clippit word verify` | Validate package/schema/relationships |
| `clippit word build init` | Scaffold a Word build manifest |
| `clippit word build run` | Merge DOCX sources from a manifest |
| `clippit word compare` | Compare two docs and emit tracked revisions |
| `clippit word assemble` | Fill a template with XML data |
| `clippit word accept-revisions` | Accept all revisions in a document |
| `clippit word to-html` | Convert DOCX to HTML/CSS |
| `clippit word from-html` | Convert HTML/CSS to DOCX |

All commands support `--format json|text` and `--quiet`.

## `word verify`

```text
clippit word verify <input.docx|-> [--office-version <version>] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--office-version` | Open XML schema version (default: `Microsoft365`). |

Readable-but-invalid documents return verify JSON and exit code `4`.

```bash
clippit word verify document.docx --format json
```

```json
{"input":"/work/document.docx","officeVersion":"Microsoft365","valid":true,"diagnostics":[]}
```

## `word build init`

Scaffold a `clippit-word-build.json` manifest with example section/file entries.

```text
clippit word build init [--output <manifest.json|->] [--force] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Manifest output path (default: `clippit-word-build.json`). Use `-` for stdout. |
| `--force` | Overwrite an existing manifest file. |

```bash
clippit word build init --output word-build.json
```

## `word build run`

Merge `.docx` sources listed in a manifest via `DocumentBuilder`.

```text
clippit word build run <manifest.json|-> [--output <file.docx|->] [--force] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Output DOCX path (default: manifest `output`, otherwise `merged.docx`). Use `-` for stdout. |
| `--force` | Overwrite an existing output file. |

Deck entries may use string shorthand (`"[Section]"`, `"chapter1.docx"`) or object form for
options like `start`, `count`, `keepSections`, and
`discardHeadersAndFootersInKeptSections`.

```bash
clippit word build run word-build.json --output merged.docx --format json
```

```json
{"output":"/work/merged.docx","outputSize":59321,"entryCount":2,"entries":[{"section":"Part 1"},{"file":"chapter1.docx","elements":12}]}
```

## `word compare`

Compare two `.docx` files and write a tracked-revisions result document.

```text
clippit word compare <source.docx|-> <revised.docx|-> [--output <file.docx|->] [--author <text>] [--date-time <text>] [--case-insensitive] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Output compared doc path (default: `<source>-compared.docx`). Use `-` for stdout. |
| `--author` | Author for generated revisions. |
| `--date-time` | Date/time for generated revisions. |
| `--case-insensitive` | Ignore case when comparing words. |

```bash
clippit word compare before.docx after.docx --output compared.docx --format json
```

```json
{"source":"/work/before.docx","revised":"/work/after.docx","output":"/work/compared.docx","outputSize":59321,"revisions":8,"authorForRevisions":"Jane Doe","dateTimeForRevisions":"2026-01-01T00:00:00Z","caseInsensitive":false}
```

## `word assemble`

Assemble a `.docx` from template + XML data.

```text
clippit word assemble <template.docx|-> <data.xml|-> [--output <file.docx|->] [--force] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Output assembled doc path (default: `<template>-assembled.docx`). Use `-` for stdout. |
| `--force` | Overwrite existing output file. |

Only one input can use stdin at a time. Output must not overwrite template or input XML.

```bash
clippit word assemble template.docx data.xml --output assembled.docx --format json
```

```json
{"template":"/work/template.docx","data":"/work/data.xml","output":"/work/template-assembled.docx","outputSize":42130,"templateError":false}
```

## `word accept-revisions`

Accept all tracked revisions and write a cleaned `.docx`.

```text
clippit word accept-revisions <input.docx|-> [--output <file.docx|->] [--force] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Output path (default: `<input>-accepted.docx`). Use `-` for stdout. |
| `--force` | Overwrite existing output file. |

```bash
clippit word accept-revisions draft.docx --output clean.docx --format json
```

```json
{"input":"/work/draft.docx","output":"/work/draft-accepted.docx","outputSize":42130}
```

## `word to-html`

Convert `.docx` to HTML/CSS.

```text
clippit word to-html <input.docx|-> [--output <file.html|->] [--page-title <text>] [--additional-css <css>] [--css-prefix <prefix>] [--inline-images] [--no-fabricate-css] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Output HTML path (default: `<input>.html`). Use `-` for stdout. |
| `--page-title` | HTML `<title>` text. |
| `--additional-css` | Extra CSS injected into generated `<style>`. |
| `--css-prefix` | Prefix for generated CSS classes (default: `pt-`). |
| `--inline-images` | Embed images as base64 data URIs. |
| `--no-fabricate-css` | Use inline styles instead of generated CSS classes. |

```bash
clippit word to-html report.docx --inline-images --output report.html --format json
```

```json
{"input":"/work/report.docx","output":"/work/report.html","outputSize":28473}
```

## `word from-html`

Convert HTML/CSS to `.docx`.

```text
clippit word from-html <input.html|-> [--output <file.docx|->] [--css <file>] [--default-css <file>] [--user-css <css>] [--base-uri <uri>] [--major-font <name>] [--minor-font <name>] [--font-size <pt>] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Output DOCX path (default: `<input>.docx`). Use `-` for stdout. |
| `--css`, `-c` | External author CSS file. |
| `--default-css` | Override built-in default CSS with a file. |
| `--user-css` | Additional CSS rules. |
| `--base-uri` | Base URI for resolving relative image `src` paths. |
| `--major-font` | Theme major/heading font (default: `Calibri Light`). |
| `--minor-font` | Theme minor/body font (default: `Times New Roman`). |
| `--font-size` | Default font size in points (default: `12`). |

```bash
clippit word from-html article.html --css styles.css --output article.docx --format json
```

```json
{"input":"/work/article.html","output":"/work/article.docx","outputSize":45021}
```
