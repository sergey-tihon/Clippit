# AGENTS.md

Guidelines for AI coding agents working in the Clippit repository.
Clippit is a .NET library providing OpenXml PowerTools for Word, Excel, and PowerPoint.

## Build, Lint, and Test

### Quick Reference

| Task                | Command                                                                        |
| ------------------- | ------------------------------------------------------------------------------ |
| Full build pipeline | `./build.sh` (Unix) or `.\build.cmd` (Windows)                                 |
| Build only          | `dotnet build Clippit.slnx -c Release`                                         |
| Restore tools       | `dotnet tool restore`                                                          |
| Lint check          | `dotnet csharpier check .`                                                     |
| Lint fix            | `dotnet csharpier format .`                                                    |
| Test all            | Run `dotnet test --project Clippit.Tests/` separately with `-f net10.0` and `-f net11.0` |
| Single test         | `dotnet test --project Clippit.Tests/ -f net11.0 --treenode-filter "/*/*/*/MethodName**"` |
| Tests in class      | `dotnet test --project Clippit.Tests/ -f net11.0 --treenode-filter "/*/*/ClassName/**"` |
| Check outdated deps | `dotnet outdated`                                                              |

### Build Pipeline

The full build script (`build.sh`/`build.cmd`) calls `dotnet fsi build.fsx -- -p build`,
which runs these stages in order:

1. `dotnet tool restore` and `dotnet restore`
2. `dotnet csharpier check .` (formatting check — build fails if violated)
3. `dotnet clean` and clean `bin/`
4. Generate `Clippit/Properties/AssemblyInfo.g.cs` (version from `CHANGELOG.md`)
5. `dotnet build Clippit.slnx -c Release` (CI builds one test TFM per matrix job)
6. Run `dotnet test --solution Clippit.slnx -f net10.0`, then run it with `-f net11.0`.
7. `dotnet pack Clippit/Clippit.csproj -o bin/`

Always run `dotnet csharpier format .` before committing. A pre-commit hook (Husky.NET)
runs CSharpier on staged `.cs` files automatically.

### Test Runner

Tests use **TUnit** with the `Microsoft.Testing.Platform` runner (configured in `global.json`).
The library and test projects target `net10.0` and `net11.0`; the CLI targets `net11.0` only. Run tests
for each framework separately to avoid collisions in the shared temp directory. The `--treenode-filter`
flag uses TUnit's tree filter syntax — `/*/*/*/MethodName**` matches any test method starting with
`MethodName` across all assemblies/namespaces/classes.

## Project Structure

```
Clippit/                    Main library (targets net10.0 and net11.0)
  Word/                     Word/DOCX: DocumentBuilder, DocumentAssembler, WmlComparer, HtmlConverter
    Assembler/              DocumentAssembler helpers (templates, extensions)
  PowerPoint/               PPTX: PresentationBuilder, Fluent API
  Excel/                    XLSX: SpreadsheetWriter, SmlDataRetriever
  Html/                     HTML-to-WML conversion (CssParser, CssApplier)
  Comparer/                 WmlComparer (partial classes split across ~26 files by concern)
  Core/                     PowerToolsBlock, StronglyTypedBlock
  Internal/                 ColorParser, TextReplacer, Relationships
  Properties/               AssemblyInfo.cs (InternalsVisibleTo), AssemblyInfo.g.cs (generated)
Clippit.Cli/                CLI tool (targets net11.0)
  Commands/Pptx/Split/      `pptx split` command and result payload
  Commands/Pptx/Build/      `pptx build init` / `pptx build run`, manifests, sections
  Commands/Pptx/Verify/     `pptx verify` validation command and result payload
  Commands/Version/         `version` command and root `--version` payload
  Infrastructure/           Output format, structured errors, exit codes
Clippit.Tests/              Test project (targets net10.0 and net11.0)
  Word/ PowerPoint/ Excel/ Html/ Common/    Mirror library structure
  Cli/                                      CLI integration and AOT smoke tests
  */Samples/                                Sample/integration tests
TestFiles/                  Test data (.docx, .pptx, .xlsx)
docs/schemas/               JSON schemas for CLI manifests/results
Directory.Build.props       Shared MSBuild props (nullable, implicit usings, lang version)
.editorconfig               Code style rules
```

## Code Style Guidelines

### Indentation and Formatting

- 4 spaces for C#; 2 spaces for XML/JSON/scripts
- Max line length: 120 characters
- CSharpier handles all formatting — do not fight it
- UTF-8 encoding (with BOM for `.cs` files)
- Always insert final newline; trim trailing whitespace

### Imports / Usings

- `System.*` usings first, then all others
- No blank lines between using groups
- Implicit usings are enabled — do not add explicit usings for `System`,
  `System.Collections.Generic`, `System.Linq`, `System.Threading.Tasks`, etc.

### Type Declarations

- Prefer `var` for all local variable declarations
- Use language keywords (`int`, `string`, `bool`) over framework types (`Int32`, `String`)
- Nullable reference types are enabled — annotate nullability, avoid `null!` suppression

### Naming Conventions

| Symbol                     | Convention       | Example                           |
| -------------------------- | ---------------- | --------------------------------- |
| Types, methods, properties | PascalCase       | `DocumentBuilder`, `GetMetrics()` |
| Constants, enum members    | PascalCase       | `MaxRetries`, `Landscape`         |
| Local functions            | PascalCase       | `ProcessElement()`                |
| Private static fields      | `s_` + camelCase | `s_tempDir`, `s_maxId`            |
| Private instance fields    | `_` + camelCase  | `_package`, `_validator`          |
| Locals, parameters         | camelCase        | `sourceDoc`, `partUri`            |

Avoid `this.` qualification for field/property/method access.

### Braces and Control Flow

- Allman style: newline before every opening brace
- Always use braces for `if`/`else`/`for`/`foreach`/`while`/`using` blocks
- Expression bodies for properties, indexers, and accessors
- Block bodies for methods, constructors, and operators

### Modern C# Patterns

Prefer modern C# idioms wherever applicable:

- Object and collection initializers
- Null coalescing (`??`) and null propagation (`?.`)
- Pattern matching (`is`, `switch` expressions) over `as` + null check
- `is null` / `is not null` over `== null` / `!= null`
- Primary constructors for simple exception/record types
- Collection expressions (`[]`) for empty collections
- `using` declarations (without braces) where scope allows

### Error Handling

- `AnalysisMode` is `All` but warnings are **not** treated as errors
- Use project-specific exceptions: `OpenXmlPowerToolsException`,
  `PowerToolsDocumentException`, `InvalidOpenXmlDocumentException`,
  `DocumentBuilderException`, `PresentationBuilderException`
- Guard clauses: prefer `ArgumentNullException.ThrowIfNull(param)` for null checks
- No logging framework — the library does not log; tests use `Console.WriteLine`

### Namespaces

- File-scoped namespaces preferred for new code (`namespace Clippit.Word;`)
- Some legacy files use block-scoped namespaces — either style is acceptable

## Testing Guidelines

### Framework and Assertions

- Testing framework: **TUnit** — use `[Test]` and `[Arguments]` attributes
- Assertions are async and fluent:

    ```csharp
    await Assert.That(errors).IsEmpty();
    await Assert.That(value).IsEqualTo(expected);
    await Assert.That(collection).HasCount(5);
    await Assert.That(() => action).Throws<OpenXmlPowerToolsException>();
    ```

### Test Structure

- All test classes inherit from `TestsBase` (provides `TempDir`, `Validate()`, helpers)
- Test data lives in `TestFiles/` — reference via `new DirectoryInfo("../../../../TestFiles/")`
- Test output goes to `temp/` directory (created lazily by `TestsBase.TempDir`)
- CLI tests use `CliTestRunner.CreateTempDirectory()` under the OS temp directory instead of `TestsBase.TempDir`
  because `TestsBase` may wipe repo `temp/` during parallel test runs.
- Organize tests to mirror library structure (`Word/`, `PowerPoint/`, etc.)
- Name tests with a prefix code and descriptive name: `DB001_DocumentBuilderKeepSections`

### Test Requirements

- Always validate OpenXml output with `Validate()` from `TestsBase`
- Parameterized tests: stack multiple `[Arguments]` on a single `[Test]` method
- Synchronous tests return `void`; async tests return `Task`
- Bug fix PRs must include test case(s) that reproduce the issue and verify the fix

## Key Architecture Notes

- **Partial classes**: `WmlComparer` is split across ~26 files by concern
  (e.g., `WmlComparer.Public.Methods.Compare.cs`, `WmlComparer.Private.Methods.Hashing.cs`)
- **Extension methods**: `PtOpenXmlExtensions`, `PtExtensions` extend OpenXml SDK types
- **CLI output discipline**: success payloads go to stdout (compact JSON when `--format json` or stdout is piped,
  text otherwise); command errors always go to stderr as compact JSON with a symbolic `code`.
- **CLI validation**: `pptx verify` combines `OpenXmlValidator` schema checks with `RelationshipValidator` dangling
  relationship checks; invalid-but-readable decks return a result on stdout and exit code `4`.
- **XNamespace constants**: `W`, `WP`, `M`, `MC`, etc. in `PtOpenXmlUtil.cs` define
  all XML namespace constants used throughout the library
- **PPTX memory model**: large-deck workflows should open source presentations with
  `OpenSettings { AutoSave = false }`, copy slides through `FluentPresentationBuilder`, then release per-slide XML
  caches. For source slides call `RemoveAnnotations<XDocument>()` and `UnloadRootElement()` after copying. For
  destination slides that were mutated directly, call `PutXDocument()` and `RemoveAnnotations<XDocument>()`. Dispose
  in this order: fluent builder, presentation document, stream; save sections after builder disposal and before
  disposing the destination document.
- **Versioning**: version is derived from `CHANGELOG.md` by the build script
- **No new dependencies** without clear justification — the project values minimal deps
