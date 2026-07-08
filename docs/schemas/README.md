# Clippit JSON Schemas

These schemas describe the JSON wire formats produced and consumed by the
Clippit CLI.

The deck manifest schema is embedded in generated manifests as `$schema`, so
editors (VS Code, JetBrains, etc.) can validate and autocomplete authored
manifest files. Result schemas are not embedded in stdout payloads; they are
published for documentation, integration validation, and contract tests.

| File                                                   | What it describes                                                                                                                                                           |
| ------------------------------------------------------ | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| [`deck-manifest.v1.json`](./deck-manifest.v1.json)     | Input manifest consumed by `clippit pptx build run`                                                                                                                         |
| [`word-build-manifest.v1.json`](./word-build-manifest.v1.json) | Input manifest consumed by `clippit word build run`                                                                                                               |
| [`split-result.v1.json`](./split-result.v1.json)       | Stdout payload of `clippit pptx split` (JSON mode)                                                                                                                          |
| [`build-result.v1.json`](./build-result.v1.json)       | Stdout payload of `clippit pptx build run` (JSON mode)                                                                                                                      |
| [`word-build-result.v1.json`](./word-build-result.v1.json) | Stdout payload of `clippit word build run` (JSON mode)                                                                                                                  |
| [`verify-result.v1.json`](./verify-result.v1.json)     | Stdout payload of `clippit pptx verify`, `clippit word verify`, `clippit excel verify` (JSON mode)                                                                          |
| [`compare-result.v1.json`](./compare-result.v1.json)   | Stdout payload of `clippit word compare` (JSON mode)                                                                                                                        |
| [`assemble-result.v1.json`](./assemble-result.v1.json) | Stdout payload of `clippit word assemble` (JSON mode)                                                                                                                       |
| [`consolidate-result.v1.json`](./consolidate-result.v1.json) | Stdout payload of `clippit word consolidate` (JSON mode)                                                                                                            |
| [`convert-result.v1.json`](./convert-result.v1.json)   | Stdout payload of `clippit word to-html`, `clippit word from-html`, `clippit word accept-revisions`, `clippit word simplify-markup`, or `clippit excel to-html` (JSON mode) |

## Output discipline

- **Success** → JSON to stdout (compact, one line) when `--format json` is
  passed or stdout is piped. Human-readable text otherwise. `--quiet`
  suppresses the success payload entirely; the exit code is the source of
  truth.
- **Command execution errors** → after arguments have parsed successfully,
  errors are emitted as a single compact JSON object on stderr regardless of
  `--format` or `--quiet`. The shape is:

  ```json
  { "error": "<human-readable message>", "code": "<symbolic code>" }
  ```

  Parser/help errors come from System.CommandLine and may include usage text.

  `pptx verify`, `word verify`, and `excel verify` are intentionally different
  for validation failures: an invalid but readable file is a successful
  verification result with `valid: false` on stdout and process exit code `4`.
  Operational failures (missing file, IO errors, unexpected exceptions) still
  use the stderr error shape above.

  `word to-html`, `word from-html`, and `excel to-html` emit their conversion result on stdout when
  converted content is written to a file. When `--output -` is used, stdout is
  reserved for the converted HTML/DOCX stream and the success payload is
  suppressed.

  `word compare` and `word consolidate` follow the same pattern: when
  `--output -` is used, stdout carries the binary DOCX and the JSON result is
  suppressed.

## Exit codes

| Code | Symbolic            | Meaning                                               |
| ---- | ------------------- | ----------------------------------------------------- |
| 0    | —                   | Success                                               |
| 1    | `INTERNAL_ERROR`    | Unexpected / unclassified failure                     |
| 2    | `INVALID_ARGUMENTS` | CLI argument or manifest validation failed            |
| 3    | `FILE_NOT_FOUND`    | A referenced input file does not exist                |
| 4    | `INVALID_FORMAT`    | A file exists but is not a valid OpenXml / JSON input |
| 5    | `OUTPUT_ERROR`      | Could not write output (permission, collision, etc.)  |

## Diagnostics

`pptx verify`, `word verify`, and `excel verify` diagnostics use a unified
`kind` field. SDK diagnostics map from OpenXml SDK `ValidationErrorType`
values; Clippit-specific diagnostics use custom kind values.

Current diagnostic kinds:

| Kind                   | Source                              |
| ---------------------- | ----------------------------------- |
| `schema`               | OpenXml SDK schema validation       |
| `semantic`             | OpenXml SDK semantic validation     |
| `package`              | OpenXml package/readability checks  |
| `markupCompatibility`  | OpenXml SDK compatibility checks    |
| `relationship`         | Clippit dangling relationship check |
| `presentation.section` | Clippit PPTX section validation     |
| `unknown`              | Future/unknown SDK validation type  |

SDK diagnostics may report `element` as `{namespace}localName`. Relationship
diagnostics report element and attribute local names plus the dangling
`relationshipId`.
