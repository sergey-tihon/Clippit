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

### SpreadsheetWriter Sample

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