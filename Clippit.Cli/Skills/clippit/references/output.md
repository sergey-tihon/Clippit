# Clippit CLI output contract

Use this when scripting or when another tool must consume results.

## Streams

- Successful command payloads are written to stdout.
- Command execution errors are written to stderr as compact JSON with `error` and `code`.
- Parser/help errors are produced by System.CommandLine and may include usage text.

## Formats

Use JSON for automation:

```bash
clippit pptx verify deck.pptx --format json
clippit word compare before.docx after.docx --output compared.docx --format json
```

Use `--quiet` when only the exit code matters:

```bash
clippit word verify document.docx --quiet
```

Do not parse human text output. Re-run with `--format json` instead.

## Exit code categories

- `0`: success
- `1`: internal error
- `2`: invalid arguments
- `3`: file not found
- `4`: invalid Office/OpenXml/JSON format
- `5`: output/write error

For validation commands, an invalid-but-readable Office file can return diagnostics and a non-zero exit code. Report diagnostics to the user.

## Safe automation pattern

```bash
if clippit pptx verify deck.pptx --format json >verify.json 2>error.json; then
  echo "valid"
else
  echo "invalid or failed; inspect verify.json and error.json"
fi
```

## Binary piping

Some commands support `-` for stdin or stdout. Only use binary stdout when explicitly needed, and avoid mixing text summaries with binary streams. Prefer file outputs for agent workflows.
