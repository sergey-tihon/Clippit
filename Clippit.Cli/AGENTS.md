# AGENTS.md (Clippit.Cli)

Guidelines for AI coding agents working in the CLI project.
Applies to files under `Clippit.Cli/` and CLI-specific tests under `Clippit.Tests/Cli/`.

## Architecture

- Keep command files thin: argument/option definitions, help text, and action binding only.
- Put execution/business logic in command services (for example `PptxSplitService`, `PptxBuildRunService`, `PptxVerifyService`).
- Use `CommandRunner.Execute` as the single top-level exception boundary for command actions.

## Output Contract

- Success payloads go to stdout.
- Command execution errors go to stderr as compact JSON (`{ error, code }`).
- Parser/help errors are produced by System.CommandLine and may include usage text.
- Respect `--format json|text` and `--quiet` consistently for every command.
- For binary stdout flows (for example `--output -`), suppress success summaries to avoid stream corruption.

## Exit Codes and Error Codes

- Keep exit code mappings in `Infrastructure/ExitCodes.cs` and symbolic codes in `ErrorCodes`.
- Prefer throwing `CliException` for user-facing/domain failures.
- Treat malformed OpenXml and invalid JSON inputs as `INVALID_FORMAT`.

## JSON / AOT

- Add every new CLI JSON DTO to `CliJsonContext` source-generation attributes.
- Do not rely on runtime reflection for serialization.
- Keep payload property names and schema contracts stable unless intentionally changed.

## Schemas and Contract Tests

- Keep result schemas in `docs/schemas/` in sync with command outputs.
- For payload shape changes, update both schema and CLI integration tests.

## Publishing NativeAOT Binaries

GitHub Actions builds and publishes release binaries/packages. Do not try to produce the full
cross-platform npm release locally unless explicitly asked.

For local smoke testing, the `publish` pipeline in `build.fsx` can produce self-contained NativeAOT
binaries and npm packages for a RID subset. Set `CLIPPIT_PUBLISH_RIDS` to a comma-separated list.

### macOS (osx-arm64, osx-x64)

These build natively on macOS arm64:

```bash
CLIPPIT_PUBLISH_RIDS=osx-arm64,osx-x64 dotnet fsi build.fsx -- -p publish
```

### Linux/Windows

NativeAOT cross-compilation is platform/toolchain dependent. Use the CI workflow for linux-x64 and
win-x64 release artifacts.

Linux contributors can smoke-test linux-x64 locally on a Linux host with the NativeAOT toolchain
installed:

```bash
CLIPPIT_PUBLISH_RIDS=linux-x64 dotnet fsi build.fsx -- -p publish
```

Windows contributors can smoke-test win-x64 locally from a developer shell:

```powershell
$env:CLIPPIT_PUBLISH_RIDS = "win-x64"
dotnet fsi build.fsx -- -p publish
```

## Tests

- Prefer focused integration tests grouped by command family:
  - `VersionTests`
  - `PptxSplitTests`
  - `PptxBuildTests`
  - `PptxVerifyTests`
- Keep CLI tests deterministic and avoid depending on terminal formatting.
