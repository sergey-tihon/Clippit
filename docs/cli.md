# Clippit CLI

Clippit CLI is a scriptable command-line interface for PowerPoint (`pptx`), Word (`word`), and Excel (`excel`) document workflows. This documentation is split into focused pages by command area.

## CLI pages

| Page | Scope |
| --- | --- |
| [PowerPoint CLI (`pptx`)](cli-pptx.md) | `pptx split`, `pptx build init`, `pptx build run`, `pptx verify` |
| [Word CLI (`word`)](cli-word.md) | `word verify`, `word build init`, `word build run`, `word compare`, `word consolidate`, `word assemble`, `word accept-revisions`, `word simplify-markup`, `word to-html`, `word from-html` |
| [Excel CLI (`excel`)](cli-excel.md) | `excel verify`, `excel to-html`, `excel create` |

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

## Workspace skills

Clippit can install local Agent Skills files for coding assistants that support
workspace skills.

Run the installer from the workspace where you want the skill files to live:

```bash
clippit install --skills          # installs .agents/skills/clippit
clippit install --skills=claude   # installs .claude/skills/clippit
clippit install --skills=all      # installs both targets
```

Running the command again replaces the installed skill files with the current
bundled versions. Use `--dry-run` to preview the paths that would be written.

```bash
clippit install --skills=all --dry-run
```

For full command options and payload contracts, use these docs,
`clippit <command> --help`, and the schemas listed below.

## Supported platforms

- **Continuously tested:** Windows, Linux
- **Runtime requirement:** .NET 10 runtime/toolchain for the `dotnet tool`; the npm package bundles native binaries for supported platforms
- **Package managers:** `dotnet tool` (NuGet) and `npm`

## Output contract

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
cat word-build.json | clippit word build run - --output - > merged.docx
cat article.html | clippit word from-html - --output - > article.docx
cat spreadsheet.xlsx | clippit excel verify - --format json
```

## `install --skills`

Installs bundled Clippit workspace skills.

```text
clippit install --skills[=agents|claude|all] [--dry-run] [--format json|text] [--quiet]
```

Targets:

| Target | Install path |
| --- | --- |
| `agents` (default) | `.agents/skills/clippit/SKILL.md` |
| `claude` | `.claude/skills/clippit/SKILL.md` |
| `all` | both target directories |

Text output lists installed skill paths. JSON output reports the installed targets,
paths, and CLI version:

```json
{"installed":[{"target":"agents","path":".agents/skills/clippit/SKILL.md"}],"version":"0.8.0"}
```

`--dry-run` prints the planned `SKILL.md` paths instead of writing files. In JSON
mode it returns:

```json
{"paths":[".agents/skills/clippit/SKILL.md"]}
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
| Word build manifest | `https://sergey-tihon.github.io/Clippit/schemas/word-build-manifest.v1.json` |
| Excel workbook definition | `https://sergey-tihon.github.io/Clippit/schemas/workbook-definition.v1.json` |
| `pptx split` result | `https://sergey-tihon.github.io/Clippit/schemas/split-result.v1.json` |
| `pptx build run` result | `https://sergey-tihon.github.io/Clippit/schemas/build-result.v1.json` |
| `word build run` result | `https://sergey-tihon.github.io/Clippit/schemas/word-build-result.v1.json` |
| `pptx verify`, `word verify`, `excel verify` result | `https://sergey-tihon.github.io/Clippit/schemas/verify-result.v1.json` |
| `word assemble` result | `https://sergey-tihon.github.io/Clippit/schemas/assemble-result.v1.json` |
| `word compare` result | `https://sergey-tihon.github.io/Clippit/schemas/compare-result.v1.json` |
| `word consolidate` result | `https://sergey-tihon.github.io/Clippit/schemas/consolidate-result.v1.json` |
| `excel create` result | `https://sergey-tihon.github.io/Clippit/schemas/excel-create-result.v1.json` |
| `install --skills` result | `https://sergey-tihon.github.io/Clippit/schemas/install-result.v1.json` |
| `install --skills --dry-run` plan | `https://sergey-tihon.github.io/Clippit/schemas/install-plan.v1.json` |
| `word to-html`, `word from-html`, `word accept-revisions`, `word simplify-markup`, `excel to-html` result | `https://sergey-tihon.github.io/Clippit/schemas/convert-result.v1.json` |
