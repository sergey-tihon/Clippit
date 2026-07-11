# Changelog

## [0.7.0] - 2026-07-05

- Added `excel create` command to generate an `.xlsx` workbook from a JSON workbook
  definition. Wraps `SpreadsheetWriter`/`WorkbookDfn`. Supports one positional input
  argument (`-` for stdin). Options: `--output`/`-o` (defaults to `<input>.xlsx`,
  `-` for stdout), `--force`, `--format json`, `--quiet`. The JSON result reports
  `input`, `output`, `outputSize`, and `worksheetCount`. Invalid sheet names or
  duplicate worksheet names return `INVALID_ARGUMENTS` (exit 2). Published schemas:
  `workbook-definition.v1.json` (input) and `excel-create-result.v1.json` (result)
  (#378).
- Added `word build init` and `word build run` commands for manifest-driven DOCX
  merging with `DocumentBuilder`. `word build init` scaffolds a
  `clippit-word-build.json` manifest (or writes it to stdout with `--output -`).
  `word build run` reads a manifest from a file or stdin, supports shorthand or
  object-form deck entries, writes the merged `.docx` to a file or stdout, and
  reports per-entry element counts in JSON mode. Published schemas:
  `word-build-manifest.v1.json`, `word-build-result.v1.json` (#374).
- Added `word consolidate` command to combine multiple revisions of a document into one
  file with tracked changes. Wraps `WmlComparer.Consolidate`. Supports one positional
  original argument and one or more revision arguments. Reviewer names and hex colors
  can be assigned with `--revisor` and `--color` (repeated per revision; defaults to the
  revision file name and a rotating palette). Additional options: `--author`,
  `--date-time`, `--case-insensitive`, `--no-table-consolidation`, `--output`/`-o`
  (defaults to `<original>-consolidated.docx`, `-` for stdout), `--force`,
  `--format json`, `--quiet`. Supports stdin (`-`) for the original document; revision
  files must be filesystem paths. Mismatched `--revisor`/`--color` counts or zero
  revision files return `INVALID_ARGUMENTS` (exit 2). When `--output -` is used, stdout
  is reserved for the binary DOCX stream and the JSON result is suppressed. Published
  schema: `consolidate-result.v1.json` (#379).
- Added `word simplify-markup` command to remove non-content markup from a `.docx` file.
  Wraps `MarkupSimplifier.SimplifyMarkup`. Accepts one flag per `SimplifyMarkupSettings`
  boolean field (`--accept-revisions`, `--remove-rsid-info`, `--remove-markup-for-document-comparison`,
  `--remove-comments`, `--remove-bookmarks`, `--remove-content-controls`,
  `--remove-end-and-footnotes`, `--remove-field-codes`, `--remove-go-back-bookmark`,
  `--remove-hyperlinks`, `--remove-last-rendered-page-break`, `--remove-permissions`,
  `--remove-proof`, `--remove-smart-tags`, `--remove-soft-hyphens`, `--remove-web-hidden`,
  `--replace-tabs-with-spaces`, `--normalize-xml`) plus `--all` as a convenience preset
  that enables every option. Supports `--output`/`-o` (defaults to
  `<input>-simplified.docx`, `-` for stdout), `--force`, `--format json`, and `--quiet`.
  At least one flag must be supplied or `INVALID_ARGUMENTS` (exit 2) is returned.
  The JSON result reuses the `convert-result.v1.json` shape (`input`, `output`,
  `outputSize`). Supports stdin (`-`) for input (#377).
- Added `word assemble` command to generate a `.docx` from a template and XML data.
  Wraps `DocumentAssembler.AssembleDocument`. Supports `--output`/`-o` (defaults to
  `<template>-assembled.docx`, `-` for stdout), `--force`, `--format json`, and
  `--quiet`. The JSON result reports `template`, `data`, `output`, `outputSize`,
  and `templateError`. Supports stdin (`-`) for one input. Published schema:
  `assemble-result.v1.json` (#375).
- Added `word accept-revisions` command to accept all tracked revisions in a `.docx` file.
  Wraps `RevisionAccepter.AcceptRevisions`. Supports `--output`/`-o` (defaults to
  `<input>-accepted.docx`, `-` for stdout), `--force`, `--format json`, and `--quiet`.
  The JSON result reports `input`, `output`, and `outputSize` (same shape as
  `convert-result.v1.json`). Supports stdin (`-`) for input (#376).

## [0.6.0] - 2026-06-29

- Added `word compare` command to compare two `.docx` files with `WmlComparer` and
  write a tracked-revision result document. Supports `--output`/`-o` (defaults to
  `<source>-compared.docx`, `-` for stdout), `--author` and `--date-time` for
  generated revision metadata, and `--case-insensitive`. The JSON result reports
  `source`, `revised`, `output`, `outputSize`, `revisions`, `authorForRevisions`,
  `dateTimeForRevisions`, and `caseInsensitive`. Supports stdin (`-`) for one input,
  `--format json`, and `--quiet`. Published schema: `compare-result.v1.json` (#365).
- feat(npm): add Windows ARM64 native binary support — new npm platform package
  `@sergey-tihon/clippit-bin-win32-arm64` built on GitHub Actions
  `windows-11-arm` runners (#363)

## [0.5.0] - 2026-06-29

- feat(npm): add Linux ARM64 native binary support — new npm platform package
  `@sergey-tihon/clippit-bin-linux-arm64` built on GitHub Actions
  `ubuntu-24.04-arm` runners (#360)
- chore(deps): bump actions/checkout from 6 to 7 (#347)

## [0.4.1] - 2026-06-23

- chore(deps): update SkiaSharp 3.119.4 → 4.148.0 with migration fix (#346)
- chore(deps): update System.CommandLine 3.0.0-preview.4 → preview.5 (#346)
- chore(deps): update dotnet-outdated-tool 4.7.1 → 4.8.1 (#346)
- chore(deps): update csharpier 1.2.6 → 1.3.0 (#346)
- chore(deps): update Microsoft.NET.Test.Sdk 18.6.0 → 18.7.0 (#346)

## [0.4.0] - 2026-06-20

- refactor: replace SixLabors.ImageSharp with SkiaSharp for image processing in
  `word to-html` inline image encoding and CLI image handler (#341)
- fix(cli): use `FileMode.Create` instead of `FileMode.OpenOrCreate` in
  `DefaultImageHandler` to avoid corrupted images on re-encode (#341)
- fix(cli): return `null` instead of `null!` in `CreateInlineImage` failure
  paths when SkiaSharp encode fails (#341)
- chore(deps): remove `SixLabors.ImageSharp.Drawing`, add `SkiaSharp 3.119.4` (#341)

## [0.3.0] - 2026-06-15

- Added `word to-html` command to convert `.docx` files to HTML/CSS with
  support for embedded or external images, custom CSS injection, page title
  control, and CSS class prefix/fabrication options.
- Added `word from-html` command to convert HTML/CSS documents to `.docx`
  files with CSS extraction from `<style>` elements or external files, font
  configuration, and image base URI resolution.
- Both new commands support stdin/stdout pipelines (`-`), stable JSON
  output (`--format json`), and quiet mode (`--quiet`).
- Added `excel to-html` command to convert Excel spreadsheet sheets, ranges,
  or named tables to HTML/CSS.
  - `--sheet` selects a worksheet by name (defaults to the first sheet).
  - `--range` restricts conversion to a cell range (e.g. `A1:D10`); requires `--sheet`.
  - `--table` converts a named Excel table and emits a `<caption>` with the table name;
    cannot be combined with `--sheet` or `--range`.
  - `--page-title`, `--additional-css`, `--css-prefix`, and `--no-fabricate-css` mirror
    the options available on `word to-html`.
  - Supports stdin/stdout pipelines (`-`), `--format json`, and `--quiet`.
  - Cell formatting (font, fill, border, alignment) is translated to CSS.

## [0.2.0] - 2026-06-05

- Added `word verify` command to validate `.docx` files.
- Added `excel verify` command to validate `.xlsx` files.

## [0.1.3] - 2026-06-03

Add README and CLI usage documentation to npm and NuGet packages.

## [0.1.2] - 2026-06-02

Patch release for the next CLI package publish.

## [0.1.1] - 2026-06-01

Patch release focused on NativeAOT packaging/runtime stability.

- Fixed NativeAOT runtime failures in PowerPoint command paths (`pptx split`,
  `pptx build`) caused by trim-sensitive type loading.
- Fixed NativeAOT linker descriptor wiring and rooted required Fluent/
  PowerPoint exception types for AOT publish.
- Fixed NativeAOT executable naming in publish output (stable `clippit` /
  `clippit.exe` output for npm packaging).
- Fixed npm platform package naming and publish workflow for scoped packages:
  `@sergey-tihon/clippit-bin-win32-x64`,
  `@sergey-tihon/clippit-bin-darwin-x64`,
  `@sergey-tihon/clippit-bin-darwin-arm64`,
  `@sergey-tihon/clippit-bin-linux-x64`.
- Fixed `pack-npm` path resolution for scoped package names by separating npm
  package name from local package directory name.
- Fixed npm publish workflow flags for first publish + provenance
  (`--access public`).

## [0.1.0] - 2026-05-23

Initial CLI release for Clippit.

- PowerPoint commands:
  - `pptx split`: split a `.pptx` into individual single-slide presentations.
  - `pptx split --slides`: extract selected slides using `N`, `N-M`, 1-based, inclusive, comma-separated syntax.
  - `pptx split --manifest`: generate a `pptx build run` compatible manifest alongside split slides, preserving source PPTX sections when present.
  - `pptx build init`: scaffold a deck manifest.
  - `pptx build run`: build a `.pptx` from a deck manifest, including per-entry slide-count reporting.
  - `pptx verify`: validate PPTX package, OpenXml schema, relationships, markup compatibility, and presentation sections with structured diagnostics.
- Input/output features:
  - stdin support via `-` for `pptx split`, `pptx build run`, and `pptx verify`.
  - stdout support via `-` for `pptx build init` manifests and `pptx build run` binary PPTX output.
  - `--force` for commands that write files and may overwrite existing output.
  - `--quiet` / `-q` on every command to suppress success output while preserving exit codes.
  - `--format json|text` for structured JSON or human-readable text output.
- Machine-readable behavior:
  - Compact JSON success payloads on stdout.
  - Compact JSON command execution errors on stderr with stable symbolic error codes.
  - Parser/help errors use System.CommandLine usage output.
  - Structured `version` command and matching top-level `--version` output.
  - Published JSON schemas under `docs/schemas/` for deck manifests and CLI result payloads.
  - Verify result JSON uses `diagnostics` for validation findings; diagnostic entries include `kind` and optional validator `code`.
- Distribution:
  - dotnet tool package: `Clippit.Cli`.
  - NativeAOT self-contained binaries for win-x64, osx-x64, osx-arm64, and linux-x64.
  - NativeAOT local publishing is target-toolchain dependent; set `CLIPPIT_PUBLISH_RIDS` to publish a supported local subset.
  - npm packages: `clippit`, `clippit-win32-x64`, `clippit-darwin-x64`, `clippit-darwin-arm64`, `clippit-linux-x64`.
