---
uid: Tutorial.Word.DocumentAssembler.OptionalTable
---
# Optional Table Directive

Namespace: `Clippit.Word`

## Introduction

The `Table` directive in `DocumentAssembler` supports:

- `Optional` — controls behavior when the XPath `Select` expression returns no matching data.
- `HeaderRowCount` — declares how many leading rows are treated as table headers.

By default, when a `Table` directive's `Select` XPath returns no data, `DocumentAssembler` produces a template error:

> Table Select returned no data

Adding `Optional="true"` (or `Optional="1"`) suppresses this error and silently removes the table from the output document instead.

This brings `Table` into parity with `Repeat`, `Content`, and `Conditional`, which support the `Optional`
attribute.

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

### Multi-row header table

Use `HeaderRowCount` when the table has more than one header row:

```xml
<# <Table Select="Orders/Order" HeaderRowCount="2" /> #>
```

With this setting, the first two `w:tr` rows are preserved as headers and the third row is used as the prototype row.

## Accepted Values

The `Optional` attribute accepts any XSD `xs:boolean` value:

| Value   | Effect   |
| ------- | -------- |
| `true`  | Optional |
| `1`     | Optional |
| `false` | Not optional (default) |
| `0`     | Not optional (default) |

Omitting the attribute entirely is equivalent to `Optional="false"`.

The `HeaderRowCount` attribute accepts an XSD `xs:positiveInteger` value. Omitting it is equivalent to `HeaderRowCount="1"`.
When processing pre-existing metadata elements that bypass schema validation, parsed integer values less than `1` are treated as `1`.

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

- `Repeat` and `Conditional` directives also support `Optional="true"` for missing XPath results.
- For template testing and examples, see `DocumentAssemblerTests.cs` in the test project, particularly the `DA_Table_Optional_*` test methods.

Changes merged in: [#150](https://github.com/sergey-tihon/Clippit/pull/150)
