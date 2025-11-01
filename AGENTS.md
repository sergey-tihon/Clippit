# AGENTS.md

## Build, Lint, and Test

- Build: `./build.sh` (Unix) or `.\build.cmd` (Windows)
- Lint/format: `dotnet csharpier check .`
- Test all: `dotnet test Clippit.Tests/`
- Run a single test: `dotnet test --filter "FullyQualifiedName~TestName"`

## Code Style Guidelines

- Use 4 spaces for C# code, 2 for XML/JSON/scripts
- Place System.\* usings first; no separated import groups
- Prefer `var` for type declarations
- Naming:
  - PascalCase for types, methods, properties, constants
  - camelCase for locals, parameters
  - Static fields: `s_` prefix, camelCase
  - Instance fields: `_` prefix, camelCase
- Newline before open braces; always use braces for blocks
- Prefer modern C# features (object/collection initializers, null propagation, etc.)
- Enable nullable reference types and implicit usings
- Use language keywords (`var`, `int`, etc.) over framework types
- Prefer explicit error handling; warnings are not errors

No Cursor or Copilot rules are present.
