---
uid: Tutorial.Excel.SpreadsheetWriter
---
# Write Spreadsheet

Namespace: `Clippit.Excel`

Save `WorkbookDfn` to stream/file.

```csharp
public static class SpreadsheetWriter {
    public static void WriteTo(this WorkbookDfn workbook, Stream stream)
    {...}
}
```

#### Fixes

- Added API to save to stream.

### Cell Builder

The `Clippit.Excel.Builder` namespace provides a `Cell` static class with factory methods
for concise cell creation:

| Method | Returns | Description |
|---|---|---|
| `Cell.Headers(params string[])` | `CellDfn[]` | Bold string cells for column headings |
| `Cell.String(string, bool bold = false)` | `CellDfn` | String cell (auto-strips invalid XML chars) |
| `Cell.Number(int)` / `Cell.Number(long)` | `CellDfn` | Numeric cell |
| `Cell.Bool(bool?)` | `CellDfn` | Boolean cell |
| `Cell.Date(DateTime?)` | `CellDfn` | Date cell with `mm-dd-yy` format |

### SpreadsheetWriter Sample (Cell Builder)

```csharp
using Clippit.Excel;
using Clippit.Excel.Builder;

var wb = new WorkbookDfn
{
    Worksheets =
    [
        new WorksheetDfn
        {
            Name = "MyFirstSheet",
            TableName = "NamesAndRates",
            ColumnHeadings = Cell.Headers("Name", "Age", "Rate"),
            Rows =
            [
                new RowDfn
                {
                    Cells = [Cell.String("Eric"), Cell.Number(50), Cell.Number(45)]
                },
                new RowDfn
                {
                    Cells = [Cell.String("Bob"), Cell.Number(42), Cell.Number(78)]
                },
            ]
        }
    ]
};

using var stream = new MemoryStream();
wb.WriteTo(stream);
var bytes = stream.ToArray();
```

### SpreadsheetWriter Sample (Object Initializers)

```csharp 
var wb = new WorkbookDfn()
{
    Worksheets = new WorksheetDfn[]
    {
        new()
        {
            Name = "MyFirstSheet",
            TableName = "NamesAndRates",
            ColumnHeadings =
                new CellDfn[]
                {
                    new() { Value = "Name", Bold = true, },
                    new()
                    {
                        Value = "Age",
                        Bold = true,
                        HorizontalCellAlignment = HorizontalCellAlignment.Left,
                    },
                    new()
                    {
                        Value = "Rate",
                        Bold = true,
                        HorizontalCellAlignment = HorizontalCellAlignment.Left,
                    }
                },
            Rows = new RowDfn[]
            {
                new()
                {
                    Cells = new CellDfn[]
                    {
                        new() { CellDataType = CellDataType.String, Value = "Eric", },
                        new() { CellDataType = CellDataType.Number, Value = 50, },
                        new()
                        {
                            CellDataType = CellDataType.Number,
                            Value = (decimal)45.00,
                            FormatCode = "0.00",
                        },
                    }
                },
                new()
                {
                    Cells = new CellDfn[]
                    {
                        new() { CellDataType = CellDataType.String, Value = "Bob", },
                        new() { CellDataType = CellDataType.Number, Value = 42, },
                        new()
                        {
                            CellDataType = CellDataType.Number,
                            Value = (decimal)78.00,
                            FormatCode = "0.00",
                        },
                    }
                },
            }
        }
    }
};

using var stream = new MemoryStream();
wb.WriteTo(stream);
var bytes = stream.ToArray();
```