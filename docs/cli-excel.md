# Excel CLI (`excel`)

Use this page for Excel (`.xlsx`) command reference.

- Back to [CLI start page](cli.md)
- Shared verify diagnostics schema: [schemas/README.md](schemas/README.md)

## Commands at a glance

| Command | Purpose |
| --- | --- |
| `clippit excel verify` | Validate package/schema/relationships |
| `clippit excel to-html` | Convert workbook range/sheet/table to HTML/CSS |

All commands support `--format json|text` and `--quiet`.

## `excel verify`

Validate readability and Open XML structure of a spreadsheet.

```text
clippit excel verify <input.xlsx|-> [--office-version <version>] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--office-version` | Open XML schema version (default: `Microsoft365`). |

Readable-but-invalid spreadsheets return verify JSON and exit code `4`.

```bash
clippit excel verify spreadsheet.xlsx --format json
```

```json
{"input":"/work/spreadsheet.xlsx","officeVersion":"Microsoft365","valid":true,"diagnostics":[]}
```

## `excel to-html`

Convert a worksheet, range, or defined table to HTML/CSS.

```text
clippit excel to-html <input.xlsx|-> [--output <file.html|->] [--sheet <name>] [--range <A1:D10>] [--table <name>] [--page-title <text>] [--additional-css <css>] [--css-prefix <prefix>] [--no-fabricate-css] [--format json|text] [--quiet]
```

| Option | Description |
| --- | --- |
| `--output`, `-o` | Output HTML path (default: `<input>.html`). Use `-` for stdout. |
| `--sheet` | Sheet name. Defaults to first sheet if `--range` and `--table` are omitted. |
| `--range` | Cell/range coordinates (for example `A1:D10`). Requires `--sheet`. |
| `--table` | Defined Excel table name. Cannot be combined with `--sheet` or `--range`. |
| `--page-title` | HTML `<title>` text. |
| `--additional-css` | Extra CSS injected into generated `<style>`. |
| `--css-prefix` | Prefix for generated CSS classes (default: `pt-`). |
| `--no-fabricate-css` | Use inline styles instead of generated CSS classes. |

```bash
clippit excel to-html spreadsheet.xlsx --sheet "Q3 Data" --range "B2:F15" --output sheet.html --format json
```

```json
{"input":"/work/spreadsheet.xlsx","output":"/work/sheet.html","outputSize":18231}
```

