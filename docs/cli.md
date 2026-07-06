# Clippit CLI

Clippit CLI is a scriptable command-line interface for PowerPoint (`pptx`), Word (`word`), and Excel (`excel`) document workflows. This documentation is split into focused pages so humans and LLM agents can load only the command set they need.

## CLI pages

| Page | Scope |
| --- | --- |
| [PowerPoint CLI (`pptx`)](cli-pptx.md) | `pptx split`, `pptx build init`, `pptx build run`, `pptx verify` |
| [Word CLI (`word`)](cli-word.md) | `word verify`, `word compare`, `word assemble`, `word accept-revisions`, `word to-html`, `word from-html` |
| [Excel CLI (`excel`)](cli-excel.md) | `excel verify`, `excel to-html` |

## Installation

You can install Clippit CLI in two supported ways.

### 1. NuGet (`dotnet tool`)

```bash
dotnet tool install -g Clippit.Cli
clippit --version
```

```bash
dotnet tool update -g Clippit.Cli
dotnet tool uninstall -g Clippit.Cli
```

### 2. npm

```bash
npm install -g clippit
clippit --version
```

```bash
npm install -g clippit@latest
npm uninstall -g clippit
```

## Supported platforms

- **Continuously tested:** Windows, Linux
- **Runtime requirement:** .NET 10 runtime/toolchain
- **Package managers:** `dotnet tool` (NuGet) and `npm`

## Output contract (human + agent friendly)

Use `--format json` for automation. Text output is for interactive usage.

| Stream | Condition | Payload |
| --- | --- | --- |
| stdout | successful JSON command | command-specific result JSON |
| stdout | verify command with diagnostics (`valid: false`) | verify result JSON, exit code `4` |
| stdout | command writes binary/HTML to stdout (`--output -`) | raw stream only; summary is suppressed |
| stderr | command execution error | compact JSON error with stable `code` |
| stderr/stdout | parser/help output | System.CommandLine text output |

Common exit codes:

| Code | Meaning |
| --- | --- |
| `0` | Success |
| `1` | Internal error |
| `2` | Invalid arguments |
| `3` | File not found |
| `4` | Invalid format or validation failure |
| `5` | Output error |

## Stdin/stdout rules

- Commands that accept file input support `-` for stdin.
- For binary/HTML output, use `--output -` and redirect stdout.
- Only one file input can come from stdin at a time.

Examples:

```bash
cat deck.pptx | clippit pptx split - --output slides --format json
cat deck.json | clippit pptx build run - --output - > final.pptx
cat document.docx | clippit word verify - --format json
cat article.html | clippit word from-html - --output - > article.docx
cat spreadsheet.xlsx | clippit excel verify - --format json
```

## `version`

Prints Clippit CLI and Open XML SDK versions.

```text
clippit version [--format json|text] [--quiet]
clippit --version
```

```json
{"version":"0.1.0","openXmlSdkVersion":"4.0.0"}
```

## JSON schemas

Schemas for manifests/results are in [schemas/README.md](schemas/README.md).

| Payload | URL |
| --- | --- |
| Deck manifest | `https://sergey-tihon.github.io/Clippit/schemas/deck-manifest.v1.json` |
| `pptx split` result | `https://sergey-tihon.github.io/Clippit/schemas/split-result.v1.json` |
| `pptx build run` result | `https://sergey-tihon.github.io/Clippit/schemas/build-result.v1.json` |
| `pptx verify`, `word verify`, `excel verify` result | `https://sergey-tihon.github.io/Clippit/schemas/verify-result.v1.json` |
| `word assemble` result | `https://sergey-tihon.github.io/Clippit/schemas/assemble-result.v1.json` |
| `word compare` result | `https://sergey-tihon.github.io/Clippit/schemas/compare-result.v1.json` |
| `word to-html`, `word from-html`, `word accept-revisions`, `excel to-html` result | `https://sergey-tihon.github.io/Clippit/schemas/convert-result.v1.json` |
