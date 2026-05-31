# Clippit CLI

Clippit CLI exposes PowerPoint workflows from the Clippit library as scriptable commands. It is designed for build pipelines, local deck automation, and tools that need stable JSON output.

```bash
clippit --help
clippit version
clippit pptx split deck.pptx --output slides --manifest
clippit pptx build run slides/deck.manifest.json --output final.pptx
clippit pptx verify final.pptx
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
| stdout | `pptx build run --output -` | Binary `.pptx`; no success summary is written. |
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
```

When `pptx build run --output -` writes a `.pptx` to stdout, the success summary is suppressed automatically so the binary stream is not corrupted.

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

## JSON Schemas

Schemas for deck manifests and CLI result payloads are published in [docs/schemas](schemas/README.md). Result schemas are intended for documentation, integration tests, and downstream contract validation; result payloads do not embed schema URLs.

Canonical schema URLs:

| Payload | URL |
| ------- | --- |
| Deck manifest | `https://sergey-tihon.github.io/Clippit/schemas/deck-manifest.v1.json` |
| `pptx split` result | `https://sergey-tihon.github.io/Clippit/schemas/split-result.v1.json` |
| `pptx build run` result | `https://sergey-tihon.github.io/Clippit/schemas/build-result.v1.json` |
| `pptx verify` result | `https://sergey-tihon.github.io/Clippit/schemas/verify-result.v1.json` |
