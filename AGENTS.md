# AGENTS.md

## Build, Lint, and Test

- Build: `./build.sh` (Unix) or `.\build.cmd` (Windows)
- Lint/format: `dotnet csharpier check .` (fix with `dotnet csharpier .`)
- Test all: `dotnet test --project Clippit.Tests/`
- Run single test: `dotnet test --project Clippit.Tests/ --treenode-filter "/*/*/*/MethodName**"`
- Run tests in class: `dotnet test --project Clippit.Tests/ --treenode-filter "/*/*/ClassName/**"`

## Code Style Guidelines

- 4 spaces for C#, 2 for XML/JSON/scripts; max line length 120
- System.* usings first; no separated import groups
- Prefer `var` for type declarations
- Naming: PascalCase for types/methods/properties/constants; camelCase for locals/parameters
- Field prefixes: `s_` for private static, `_` for private instance (both camelCase)
- Newline before open braces; always use braces for control flow blocks
- Prefer modern C# (object/collection initializers, null propagation, pattern matching)
- Enable nullable reference types and implicit usings (per Directory.Build.props)
- Use language keywords (`int`, `string`) over framework types (`Int32`, `String`)
- Warnings are not treated as errors; prefer explicit error handling
- Testing framework: TUnit (use `[Test]`, `[Arguments]` attributes)
