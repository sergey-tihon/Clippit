# PowerPoint CLI (`pptx`)

Use this page for PowerPoint automation commands.

- Back to [CLI start page](cli.md)
- Schemas and diagnostic kinds: [schemas/README.md](schemas/README.md)

## Commands at a glance

| Command | Purpose |
| --- | --- |
| `clippit pptx split` | Split a deck into single-slide `.pptx` files |
| `clippit pptx build init` | Scaffold a deck manifest |
| `clippit pptx build run` | Build a final deck from a manifest |
| `clippit pptx verify` | Validate package/schema/relationships |

All commands support `--format json|text` and `--quiet`.

## `pptx split`

Split a `.pptx` file into single-slide presentations.

```text
clippit pptx split <input.pptx|-> [--output <dir>] [--slides <expr>] [--manifest] [--force] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Output directory (default: source directory, or current directory for stdin). |
| `--slides`, `-s` | 1-based slide numbers/ranges like `1,3,6-9` (default: all slides). |
| `--manifest` | Also write a `pptx build run` compatible manifest. |
| `--force` | Overwrite existing output files. |

```bash
clippit pptx split deck.pptx --output slides --manifest
```

```json
{"input":"/work/deck.pptx","outputDir":"/work/slides","manifest":"/work/slides/deck.manifest.json","count":2,"slides":[{"index":1,"file":"/work/slides/deck-001.pptx","title":"Intro"},{"index":3,"file":"/work/slides/deck-003.pptx","title":null}]}
```

## `pptx build init`

Create an empty deck manifest.

```text
clippit pptx build init [--output <manifest.json|->] [--force] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Manifest path (default: `./clippit-deck.json`). Use `-` for stdout. |
| `--force` | Overwrite existing manifest file. |

```bash
clippit pptx build init --output deck.json
```

Minimal manifest:

```json
{
  "$schema": "https://sergey-tihon.github.io/Clippit/schemas/deck-manifest.v1.json",
  "title": "Conference Deck",
  "output": "final.pptx",
  "deck": [
    { "section": "Opening" },
    "intro.pptx",
    { "file": "demo.pptx", "keepSections": true }
  ]
}
```

Manifest rules:

| Property | Required | Notes |
| --- | --- | --- |
| `title` | Yes | Output presentation title metadata. |
| `output` | Yes | Relative to manifest directory. |
| `deck` | Yes | Ordered entries; at least one entry. |
| string entry | No | `[Section Name]` means section divider; otherwise source `.pptx` path. |
| `{ "section": "Name" }` | No | Explicit section divider. |
| `{ "file": "deck.pptx" }` | No | Source deck entry; optional `masters`, `slides`, `keepSections`. |

## `pptx build run`

Build a final `.pptx` from a manifest.

```text
clippit pptx build run <manifest.json|-> [--output <file.pptx|->] [--force] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Override output path. Use `-` for binary stdout stream. |
| `--force` | Overwrite existing output file. |

```bash
clippit pptx build run deck.json --output final.pptx --force
```

```json
{"output":"/work/final.pptx","totalSlides":12,"entries":[{"section":"Opening"},{"file":"intro.pptx","slides":3},{"file":"demo.pptx","slides":9}]}
```

## `pptx verify`

Validate readability and Open XML structure of a `.pptx` file.

```text
clippit pptx verify <input.pptx|-> [--office-version <version>] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--office-version` | Open XML schema version (default: `Microsoft365`). |

If the deck is readable but has validation diagnostics, output is still emitted and process exits with code `4`.

```bash
clippit pptx verify deck.pptx --format json
```

```json
{"input":"/work/deck.pptx","officeVersion":"Microsoft365","valid":true,"diagnostics":[]}
```

