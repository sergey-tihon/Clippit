# AGENTS.md

Guidelines for AI coding agents working in the Clippit repository.
Clippit is a .NET library providing OpenXml PowerTools for Word, Excel, and PowerPoint.

## Build, Lint, and Test

### Quick Reference

| Task | Command |
|------|---------|
| Full build pipeline | `./build.sh` (Unix) or `.\build.cmd` (Windows) |
| Build only | `dotnet build Clippit.slnx -c Release` |
| Restore tools | `dotnet tool restore` |
| Lint check | `dotnet csharpier check .` |
| Lint fix | `dotnet csharpier .` |
| Test all | `dotnet test --project Clippit.Tests/` |
| Single test | `dotnet test --project Clippit.Tests/ --treenode-filter "/*/*/*/MethodName**"` |
| Tests in class | `dotnet test --project Clippit.Tests/ --treenode-filter "/*/*/ClassName/**"` |
| Check outdated deps | `dotnet outdated` |

### Build Pipeline

The full build script (`build.sh`/`build.cmd`) runs these stages in order:
1. `dotnet tool restore` and `dotnet restore`
2. `dotnet csharpier check .` (formatting check — build fails if violated)
3. `dotnet clean` and clean `bin/`
4. Generate `Clippit/Properties/AssemblyInfo.g.cs`
5. `dotnet build Clippit.slnx -c Release`
6. `dotnet test --solution Clippit.slnx`
7. `dotnet pack Clippit/Clippit.csproj -o bin/`

Always run `dotnet csharpier .` before committing. A pre-commit hook (Husky.NET)
runs CSharpier on staged `.cs` files automatically.

### Test Runner

Tests use **TUnit** with the `Microsoft.Testing.Platform` runner (configured in `global.json`).
The test project targets `net10.0` only. The `--treenode-filter` flag uses TUnit's tree
filter syntax — the pattern `/*/*/*/MethodName**` matches any test method starting with
`MethodName` across all assemblies/namespaces/classes.

## Project Structure

```
Clippit/                    Main library (targets net8.0 + net10.0)
  Word/                     Word/DOCX: DocumentBuilder, DocumentAssembler, WmlComparer, HtmlConverter
  PowerPoint/               PPTX: PresentationBuilder, Fluent API
  Excel/                    XLSX: SpreadsheetWriter, SmlDataRetriever
  Html/                     HTML-to-WML conversion
  Comparer/                 WmlComparer (partial classes split across many files)
  Core/                     PowerToolsBlock, StronglyTypedBlock
  Internal/                 ColorParser, TextReplacer, Relationships
Clippit.Tests/              Test project (targets net10.0)
  Word/ PowerPoint/ Excel/ Html/ Common/    Mirror library structure
  */Samples/                                Sample/integration tests
TestFiles/                  Test data (.docx, .pptx, .xlsx)
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

| Symbol | Convention | Example |
|--------|-----------|---------|
| Types, methods, properties | PascalCase | `DocumentBuilder`, `GetMetrics()` |
| Constants, enum members | PascalCase | `MaxRetries`, `Landscape` |
| Local functions | PascalCase | `ProcessElement()` |
| Private static fields | `s_` + camelCase | `s_tempDir`, `s_maxId` |
| Private instance fields | `_` + camelCase | `_package`, `_validator` |
| Locals, parameters | camelCase | `sourceDoc`, `partUri` |

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

- Warnings are **not** treated as errors (`TreatWarningsAsErrors: false`)
- Use project-specific exceptions: `OpenXmlPowerToolsException`,
  `PowerToolsDocumentException`, `InvalidOpenXmlDocumentException`
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
- Organize tests to mirror library structure (`Word/`, `PowerPoint/`, etc.)
- Name tests with a prefix code and descriptive name: `DB001_DocumentBuilderKeepSections`

### Test Requirements

- Always validate OpenXml output with `Validate()` from `TestsBase`
- Parameterized tests: stack multiple `[Arguments]` on a single `[Test]` method
- Synchronous tests return `void`; async tests return `Task`
- Bug fix PRs must include test case(s) that reproduce the issue and verify the fix

## Key Architecture Notes

- **Partial classes**: `WmlComparer` is split across many files by concern
  (e.g., `WmlComparer.Private.Methods.Hashing.cs`)
- **Extension methods**: `PtOpenXmlExtensions`, `PtExtensions` extend OpenXml SDK types
- **XNamespace constants**: `W`, `WP`, `M`, `MC`, etc. in `PtOpenXmlUtil.cs` define
  all XML namespace constants used throughout the library
- **No new dependencies** without clear justification — the project values minimal deps
