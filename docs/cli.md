# Clippit CLI

Clippit CLI exposes PowerPoint, Word, and Excel workflows from the Clippit library as scriptable commands. It is designed for build pipelines, local deck automation, DOCX↔HTML conversion, and tools that need stable JSON output.

```bash
clippit --help
clippit version
clippit pptx split deck.pptx --output slides --manifest
clippit pptx build run slides/deck.manifest.json --output final.pptx
clippit pptx verify final.pptx
clippit word assemble template.docx data.xml --output assembled.docx
clippit word compare before.docx after.docx --output compared.docx
clippit word accept-revisions draft.docx
clippit word verify document.docx
clippit word to-html document.docx
clippit word from-html article.html --css styles.css
clippit excel to-html spreadsheet.xlsx
clippit excel verify spreadsheet.xlsx
```

## Installation

You can install Clippit CLI either as a .NET global tool (NuGet) or as a Node.js global package (npm).

### Install from NuGet (`dotnet tool`)

```bash
dotnet tool install -g Clippit.Cli
clippit --version
```

Update / uninstall:

```bash
dotnet tool update -g Clippit.Cli
dotnet tool uninstall -g Clippit.Cli
```

### Install from npm

```bash
npm install -g clippit
clippit --version
```

Update / uninstall:

```bash
npm install -g clippit@latest
npm uninstall -g clippit
```

## Output Behavior

Every command supports `--format json|text` and `--quiet` unless it writes binary content to stdout. Text output is intended for humans; JSON output is compact and stable for automation.

Success payloads are written to stdout. Command execution errors are written to stderr as compact JSON with a stable symbolic `code`. Parser and help errors use System.CommandLine usage output.

## Automation Contract

Agents and scripts should use `--format json` when they need machine-readable command results.

| Stream | Condition | Payload |
| ------ | --------- | ------- |
| stdout | Successful JSON command | Command-specific result JSON. |
| stdout | `pptx verify` finds validation diagnostics | Verify result JSON with `valid: false`; process exits `4`. |
| stdout | `word verify` finds validation diagnostics | Verify result JSON with `valid: false`; process exits `4`. |
| stdout | `word assemble` | Result JSON (e.g. `{"template":...,"data":...,"output":...,"outputSize":...,"templateError":...}`); assembled DOCX written to `--output` path. |
| stdout | `word compare` | Result JSON (e.g. `{"source":...,"revised":...,"output":...,"revisions":...}`); compared DOCX written to `--output` path. |
| stdout | `excel verify` finds validation diagnostics | Verify result JSON with `valid: false`; process exits `4`. |
| stdout | `word to-html` / `word from-html` | Result JSON (e.g. `{"input":...,"output":...,"outputSize":...}`); converted content written to `--output` path. |
| stdout | `word accept-revisions` | Result JSON (e.g. `{"input":...,"output":...,"outputSize":...}`); cleaned DOCX written to `--output` path. |
| stdout | `excel to-html` | Result JSON (e.g. `{"input":...,"output":...,"outputSize":...}`); converted HTML written to `--output` path. |
| stdout | `pptx build run --output -` | Binary `.pptx`; no success summary is written. |
| stdout | `word assemble --output -` | Binary `.docx` streamed to stdout; no success summary is written. |
| stdout | `word to-html --output -` / `word from-html --output -` | Binary/HTML content streamed to stdout; no success summary is written. |
| stdout | `word accept-revisions --output -` | Binary `.docx` streamed to stdout; no success summary is written. |
| stdout | `excel to-html --output -` | HTML content streamed to stdout; no success summary is written. |
| stderr | Command execution error | Compact JSON error object: `{"error":"...","code":"..."}`. |
| stderr/stdout | Parser, arity, and help output | System.CommandLine text output, not JSON. |

Do not parse text output for automation. Prefer JSON schemas and the documented properties below.

Common exit codes:

| Code | Meaning |
| ---- | ------- |
| `0` | Success |
| `1` | Internal error |
| `2` | Invalid arguments |
| `3` | File not found |
| `4` | Invalid format or validation failure |
| `5` | Output error |

## Stdin And Stdout

Commands that accept files use `-` for stdin where binary or JSON streaming is supported.

```bash
cat deck.pptx | clippit pptx split - --output slides --format json
cat deck.json | clippit pptx build run - --output - > final.pptx
cat deck.pptx | clippit pptx verify - --format json
cat data.xml | clippit word assemble template.docx - --output assembled.docx --format json
cat before.docx | clippit word compare - after.docx --output compared.docx --format json
cat document.docx | clippit word verify - --format json
cat document.docx | clippit word accept-revisions - --output clean.docx --format json
cat document.docx | clippit word to-html - --inline-images --output -
cat article.html | clippit word from-html - --minor-font "Georgia" --output -
cat spreadsheet.xlsx | clippit excel verify - --format json
```

When `pptx build run --output -` writes a `.pptx` to stdout, the success summary is suppressed automatically so the binary stream is not corrupted.

When `word assemble --output -` writes a `.docx` to stdout, the success summary is also suppressed so the binary stream is not corrupted.

When `word to-html --output -` or `word from-html --output -` writes content to stdout, the success summary is also suppressed so the output stream is not corrupted.

When `word accept-revisions --output -` writes a `.docx` to stdout, the success summary is also suppressed so the binary stream is not corrupted.

Binary stdout output is buffered in memory before it is written to stdout. Prefer file output for very large decks.

## `version`

Prints the Clippit CLI version and the Open XML SDK version used by the tool.

Synopsis:

```text
clippit version [--format json|text] [--quiet]
clippit --version
```

```bash
clippit version
clippit version --format json
clippit --version
```

JSON example:

```json
{"version":"0.1.0","openXmlSdkVersion":"4.0.0"}
```

## `pptx split`

Splits a `.pptx` file into individual single-slide presentations.

Synopsis:

```text
clippit pptx split <input.pptx|-> [--output <dir>] [--slides <expr>] [--manifest] [--force] [--format json|text] [--quiet]
```

```bash
clippit pptx split deck.pptx --output slides
clippit pptx split deck.pptx --slides 1,3,6-9 --output slides
clippit pptx split deck.pptx --output slides --manifest
clippit pptx split deck.pptx --output slides --force
```

Options:

| Option | Description |
| ------ | ----------- |
| `--output`, `-o` | Output directory. Defaults to the source directory, or the current directory for stdin. |
| `--slides`, `-s` | 1-based slide numbers and inclusive ranges, for example `1,3,6-9`. Defaults to all slides. |
| `--manifest` | Also writes a `pptx build run` compatible manifest beside the split slides. |
| `--force` | Overwrite existing output files. |

The generated manifest preserves source presentation sections when present.

JSON example:

```json
{"input":"/work/deck.pptx","outputDir":"/work/slides","manifest":"/work/slides/deck.manifest.json","count":2,"slides":[{"index":1,"file":"/work/slides/deck-001.pptx","title":"Intro"},{"index":3,"file":"/work/slides/deck-003.pptx","title":null}]}
```

## `pptx build init`

Scaffolds an empty deck manifest.

Synopsis:

```text
clippit pptx build init [--output <manifest.json|->] [--force] [--format json|text] [--quiet]
```

```bash
clippit pptx build init
clippit pptx build init --output deck.json --force
clippit pptx build init --output - > deck.json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--output`, `-o` | Manifest path. Defaults to `./clippit-deck.json`. Use `-` to write JSON to stdout. |
| `--force` | Overwrite an existing manifest file. |

Deck manifests include a `$schema` URL so editors and CI jobs can validate authored input files against `docs/schemas/deck-manifest.v1.json`.

Minimal manifest example:

```json
{
  "$schema": "https://sergey-tihon.github.io/Clippit/schemas/deck-manifest.v1.json",
  "title": "Conference Deck",
  "output": "final.pptx",
  "deck": [
    { "section": "Opening" },
    "intro.pptx",
    { "file": "demo.pptx", "keepSections": true },
    { "section": "Appendix" },
    { "file": "appendix.pptx", "masters": true, "slides": true }
  ]
}
```

Manifest rules for agents:

| Property | Required | Notes |
| -------- | -------- | ----- |
| `title` | Yes | Written to output presentation core properties. |
| `output` | Yes | Relative paths resolve against the manifest file directory. |
| `deck` | Yes | Ordered entries; must contain at least one entry. |
| string entry | No | `[Section Name]` creates a section divider; any other string is a source `.pptx` path. |
| `{ "section": "Name" }` | No | Explicit section divider. |
| `{ "file": "deck.pptx" }` | No | Source `.pptx` entry. Optional booleans: `masters`, `slides`, `keepSections`. |

## `pptx build run`

Builds a final `.pptx` from a deck manifest.

Synopsis:

```text
clippit pptx build run <manifest.json|-> [--output <file.pptx|->] [--force] [--format json|text] [--quiet]
```

```bash
clippit pptx build run deck.json
clippit pptx build run deck.json --output final.pptx --force
clippit pptx build run deck.json --format json
cat deck.json | clippit pptx build run - --output - > final.pptx
```

Options:

| Option | Description |
| ------ | ----------- |
| `--output`, `-o` | Override the manifest output path. Use `-` to write the binary `.pptx` to stdout. |
| `--force` | Overwrite the output presentation if it already exists. |

Without `--force`, the command fails before replacing an existing output file.

JSON example:

```json
{"output":"/work/final.pptx","totalSlides":12,"entries":[{"section":"Opening"},{"file":"intro.pptx","slides":3},{"file":"demo.pptx","slides":9}]}
```

## `pptx verify`

Validates that a `.pptx` file is a readable and structurally correct Open XML presentation. The command checks package readability, Open XML schema errors, dangling relationships, markup compatibility issues, and Clippit presentation-section metadata.

Synopsis:

```text
clippit pptx verify <input.pptx|-> [--office-version <version>] [--format json|text] [--quiet]
```

```bash
clippit pptx verify deck.pptx
clippit pptx verify deck.pptx --format json
clippit pptx verify deck.pptx --office-version Office2021
cat deck.pptx | clippit pptx verify - --format json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--office-version` | Open XML schema version to validate against. Defaults to `Microsoft365`. |

Invalid-but-readable presentations produce a validation result on stdout and exit with code `4`. JSON results include `diagnostics`; each diagnostic has a `kind`, message, and optional validator code/location fields.

Current diagnostic kinds are documented in [schemas/README.md](schemas/README.md).

Valid JSON example:

```json
{"input":"/work/deck.pptx","officeVersion":"Microsoft365","valid":true,"diagnostics":[]}
```

Invalid JSON example:

```json
{"input":"/work/deck.pptx","officeVersion":"Microsoft365","valid":false,"diagnostics":[{"kind":"relationship","code":null,"description":"Dangling relationship target.","part":"/ppt/slides/slide1.xml","path":null,"element":null,"attribute":"id","relationshipId":"rId9"}]}
```

## `word verify`

Validates that a `.docx` file is a readable and structurally correct Open XML document. The command checks package readability, Open XML schema errors, dangling relationships, and markup compatibility issues.

Synopsis:

```text
clippit word verify <input.docx|-> [--office-version <version>] [--format json|text] [--quiet]
```

```bash
clippit word verify document.docx
clippit word verify document.docx --format json
clippit word verify document.docx --office-version Office2021
cat document.docx | clippit word verify - --format json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--office-version` | Open XML schema version to validate against. Defaults to `Microsoft365`. |

Invalid-but-readable documents produce a validation result on stdout and exit with code `4`. JSON results include `diagnostics`; each diagnostic has a `kind`, message, and optional validator code/location fields. The payload shape matches `pptx verify` and is documented in [schemas/README.md](schemas/README.md).

Valid JSON example:

```json
{"input":"/work/document.docx","officeVersion":"Microsoft365","valid":true,"diagnostics":[]}
```

## `word compare`

Compares two `.docx` files and writes a result document with tracked revisions.
The command wraps `WmlComparer.Compare`.

Synopsis:

```text
clippit word compare <source.docx|-> <revised.docx|-> [--output <file.docx|->] [--author <text>] [--date-time <text>] [--case-insensitive] [--format json|text] [--quiet]
```

```bash
clippit word compare before.docx after.docx
clippit word compare before.docx after.docx --output compared.docx --format json
clippit word compare before.docx after.docx --author "Jane Doe" --date-time 2026-01-01T00:00:00Z
cat before.docx | clippit word compare - after.docx --output compared.docx --format json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--output`, `-o` | Output path for the compared `.docx` file. Defaults to `<source>-compared.docx`. Use `-` to write binary content to stdout. |
| `--author` | Author value used for generated tracked revisions. |
| `--date-time` | Date/time value used for generated tracked revisions. |
| `--case-insensitive` | Ignore case when comparing words. |

The `authorForRevisions` and `dateTimeForRevisions` fields echo back the effective
values. When `--author`/`--date-time` are omitted they default to `Open-Xml-PowerTools`
and the current local time (round-trip `o` format). When provided, the exact strings are
echoed back without reformatting.

JSON example (for `--author "Jane Doe" --date-time 2026-01-01T00:00:00Z --output compared.docx --format json`):

```json
{"source":"/work/before.docx","revised":"/work/after.docx","output":"/work/compared.docx","outputSize":59321,"revisions":8,"authorForRevisions":"Jane Doe","dateTimeForRevisions":"2026-01-01T00:00:00Z","caseInsensitive":false}
```

## `word assemble`

Assembles a `.docx` document from a template and XML data.
The command wraps `DocumentAssembler.AssembleDocument`.

Synopsis:

```text
clippit word assemble <template.docx|-> <data.xml|-> [--output <file.docx|->] [--force] [--format json|text] [--quiet]
```

```bash
clippit word assemble template.docx data.xml
clippit word assemble template.docx data.xml --output assembled.docx --format json
clippit word assemble template.docx - --output assembled.docx
cat data.xml | clippit word assemble template.docx - --output assembled.docx --format json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--output`, `-o` | Output path for the assembled `.docx` file. Defaults to `<template>-assembled.docx`. Use `-` to write binary content to stdout. |
| `--force` | Overwrite the output file if it already exists. |

Only one input can be read from stdin at a time. The output path must not overwrite the template or XML data input file, even when `--force` is used.

JSON example:

```json
{"template":"/work/template.docx","data":"/work/data.xml","output":"/work/template-assembled.docx","outputSize":42130,"templateError":false}
```

## `word accept-revisions`

Accepts all tracked revisions in a `.docx` file and writes the cleaned document.
The command wraps `RevisionAccepter.AcceptRevisions`.

Synopsis:

```text
clippit word accept-revisions <input.docx|-> [--output <file.docx|->] [--force] [--format json|text] [--quiet]
```

```bash
clippit word accept-revisions draft.docx
clippit word accept-revisions draft.docx --output clean.docx --format json
clippit word accept-revisions draft.docx --output - > clean.docx
cat draft.docx | clippit word accept-revisions - --output clean.docx --format json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--output`, `-o` | Output path for the cleaned `.docx` file. Defaults to `<input>-accepted.docx`. Use `-` to write binary content to stdout. |
| `--force` | Overwrite the output file if it already exists. |

JSON example:

```json
{"input":"/work/draft.docx","output":"/work/draft-accepted.docx","outputSize":42130}
```

## `excel verify`

Validates that a `.xlsx` file is a readable and structurally correct Open XML spreadsheet. The command checks package readability, Open XML schema errors, dangling relationships, and markup compatibility issues.

Synopsis:

```text
clippit excel verify <input.xlsx|-> [--office-version <version>] [--format json|text] [--quiet]
```

```bash
clippit excel verify spreadsheet.xlsx
clippit excel verify spreadsheet.xlsx --format json
clippit excel verify spreadsheet.xlsx --office-version Office2021
cat spreadsheet.xlsx | clippit excel verify - --format json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--office-version` | Open XML schema version to validate against. Defaults to `Microsoft365`. |

Invalid-but-readable spreadsheets produce a validation result on stdout and exit with code `4`. JSON results include `diagnostics`; each diagnostic has a `kind`, message, and optional validator code/location fields. The payload shape matches `pptx verify` and is documented in [schemas/README.md](schemas/README.md).

Valid JSON example:

```json
{"input":"/work/spreadsheet.xlsx","officeVersion":"Microsoft365","valid":true,"diagnostics":[]}
```

## `word to-html`

Converts a `.docx` file to HTML/CSS with high-fidelity layout preservation.
Images can be embedded as base64 data URIs (`--inline-images`) or referenced as
separate files. The command wraps `WmlToHtmlConverter` from the Clippit library.

Synopsis:

```text
clippit word to-html <input.docx|-> [--output <file.html|->] [--page-title <text>] [--additional-css <css>] [--css-prefix <prefix>] [--inline-images] [--no-fabricate-css] [--format json|text] [--quiet]
```

```bash
clippit word to-html report.docx
clippit word to-html report.docx --page-title "Q3 Report" --additional-css "body { max-width: 800px; }"
clippit word to-html report.docx --inline-images --output - > report.html
cat document.docx | clippit word to-html - --inline-images --format json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--output`, `-o` | Output path for the generated `.html` file. Defaults to `<input>.html`. Use `-` to write HTML content to stdout. |
| `--page-title` | HTML page `<title>`. Defaults to the source file name. |
| `--additional-css` | Extra CSS rules injected into the generated `<style>` block. |
| `--css-prefix` | Prefix for auto-generated CSS class names (default: `pt-`). |
| `--inline-images` | Embed images as base64 data URIs instead of linking to external files. |
| `--no-fabricate-css` | Skip CSS class generation and use inline `style` attributes instead. |

JSON example:

```json
{"input":"/work/report.docx","output":"/work/report.html","outputSize":28473}
```

## `word from-html`

Converts an HTML document to a `.docx` file, restoring styles, images, and layout
as Open XML markup. CSS in the HTML `<style>` element is extracted automatically;
external CSS can also be passed with `--css`. The command wraps
`HtmlToWmlConverter` from the Clippit library.

Synopsis:

```text
clippit word from-html <input.html|-> [--output <file.docx|->] [--css <file>] [--default-css <file>] [--user-css <css>] [--base-uri <uri>] [--major-font <name>] [--minor-font <name>] [--font-size <pt>] [--format json|text] [--quiet]
```

```bash
clippit word from-html article.html
clippit word from-html article.html -c styles.css -o article.docx
clippit word from-html article.html --minor-font "Georgia" --font-size 11
cat article.html | clippit word from-html - --base-uri https://example.com/images/ --format json
```

Options:

| Option | Description |
| ------ | ----------- |
| `--output`, `-o` | Output path for the generated `.docx` file. Defaults to `<input>.docx`. Use `-` to write binary content to stdout. |
| `--css`, `-c` | Path to an external author CSS file. When omitted, CSS is extracted from the HTML `<style>` element. |
| `--default-css` | Path to a default CSS file to override the built-in default CSS. |
| `--user-css` | Additional CSS rules to apply as user overrides. |
| `--base-uri` | Base URI for resolving relative image `src` references. Defaults to the source HTML file's parent directory. |
| `--major-font` | Theme major (heading) font name (default: `Calibri Light`). |
| `--minor-font` | Theme minor (body) font name (default: `Times New Roman`). |
| `--font-size` | Default font size in points (default: `12`). |

JSON example:

```json
{"input":"/work/article.html","output":"/work/article.docx","outputSize":45021}
```

## JSON Schemas

Schemas for deck manifests and CLI result payloads are published in [docs/schemas](schemas/README.md). Result schemas are intended for documentation, integration tests, and downstream contract validation; result payloads do not embed schema URLs.

Canonical schema URLs:

| Payload | URL |
| ------- | --- |
| Deck manifest | `https://sergey-tihon.github.io/Clippit/schemas/deck-manifest.v1.json` |
| `pptx split` result | `https://sergey-tihon.github.io/Clippit/schemas/split-result.v1.json` |
| `pptx build run` result | `https://sergey-tihon.github.io/Clippit/schemas/build-result.v1.json` |
| `pptx verify`, `word verify`, `excel verify` result | `https://sergey-tihon.github.io/Clippit/schemas/verify-result.v1.json` |
| `word assemble` result | `https://sergey-tihon.github.io/Clippit/schemas/assemble-result.v1.json` |
| `word compare` result | `https://sergey-tihon.github.io/Clippit/schemas/compare-result.v1.json` |
| `word to-html`, `word from-html`, `word accept-revisions` result | `https://sergey-tihon.github.io/Clippit/schemas/convert-result.v1.json` |
