# Changelog

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
