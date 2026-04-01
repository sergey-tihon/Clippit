---
uid: Tutorial.Word.DocumentAssembler.OptionalTable
---
# Optional Table Directive

Namespace: `Clippit.Word`

## Introduction

The `Table` directive in `DocumentAssembler` supports an `Optional` attribute that controls the behavior when the XPath `Select` expression returns no matching data.

By default, when a `Table` directive's `Select` XPath returns no data, `DocumentAssembler` produces a template error:

> Table Select returned no data

Adding `Optional="true"` (or `Optional="1"`) suppresses this error and silently removes the table from the output document instead.

This brings `Table` into parity with `Repeat` and `Content`, which already support the `Optional` attribute.

## Usage

### Metadata Style

In a Word document template, use the `<# ... #>` metadata syntax:

```xml
<# <Table Select="Orders" Optional="true" /> #>
```

The table following this directive will be removed from the output when `Orders` returns no matching XML elements.

### Content Control Style

Alternatively, use a Word content control (structured document tag) with the `DocumentAssembler` directive:

```xml
<Table Select="Orders" Optional="true" />
```

Place this directive inside a content control that precedes the table in your document.

## Accepted Values

The `Optional` attribute accepts any XSD `xs:boolean` value:

| Value   | Effect   |
| ------- | -------- |
| `true`  | Optional |
| `1`     | Optional |
| `false` | Not optional (default) |
| `0`     | Not optional (default) |

Omitting the attribute entirely is equivalent to `Optional="false"`.

## Example

### XML Data (no orders)

```xml
<Data>
  <Customer>
    <Name>Cheryl</Name>
  </Customer>
</Data>
```

### Template with Optional Table

```xml
<# <Table Select="Customer/Orders/Order" Optional="true" /> #>

| Product | Quantity |
| ------- | -------- |
| ...     | ...      |
```

**Result:** The table is silently removed from the output. No template error is produced.

### Template without Optional

```xml
<# <Table Select="Customer/Orders/Order" /> #>
```

**Result:** `DocumentAssembler` sets `returnedTemplateError = true` and inserts a `[Template error: Table Select returned no data]` placeholder in the output.

## Comparison: Optional vs. Non-Optional behavior

| Scenario                          | `Optional` absent / `false` | `Optional="true"` |
| --------------------------------- | --------------------------- | ----------------- |
| XPath returns data                | Table is populated          | Table is populated |
| XPath returns no data             | Template error produced     | Table is removed silently |

## Further Reading

- `Repeat` directive also supports `Optional="true"` — see the existing directive reference for `Repeat`.
- For template testing and examples, see `DocumentAssemblerTests.cs` in the test project, particularly the `DA_Table_Optional_*` test methods.

Changes merged in: [#150](https://github.com/sergey-tihon/Clippit/pull/150)
