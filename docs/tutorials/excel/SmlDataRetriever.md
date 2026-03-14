---
uid: Tutorial.Excel.SmlDataRetriever
---
# Retrieve Spreadsheet Data

Namespace: `Clippit.Excel`

Extract sheet data, named ranges, and table data from Excel documents as XML.

```csharp
public static class SmlDataRetriever {
    public static XElement RetrieveSheet(SmlDocument smlDoc, string sheetName)
    {...}
    public static XElement RetrieveSheet(string fileName, string sheetName)
    {...}
    public static XElement RetrieveSheet(SpreadsheetDocument sDoc, string sheetName)
    {...}

    public static XElement RetrieveRange(SmlDocument smlDoc, string sheetName, string range)
    {...}
    public static XElement RetrieveRange(
        SmlDocument smlDoc, string sheetName,
        int leftColumn, int topRow, int rightColumn, int bottomRow)
    {...}

    public static XElement RetrieveTable(SmlDocument smlDoc, string sheetName, string tableName)
    {...}

    public static string[] SheetNames(SmlDocument smlDoc)
    {...}
    public static string[] TableNames(SmlDocument smlDoc)
    {...}
}
```

`SmlDataRetriever` reads Excel data and returns it as `XElement` trees. Cell values are resolved
(including shared strings), and formatting information is preserved. Each method has overloads
accepting `SmlDocument`, `string` (file path), or `SpreadsheetDocument`.

#### Key Features

- **Full sheet retrieval** -- get all data from a named worksheet
- **Range retrieval** -- extract a specific cell range by address (e.g., `"A1:D10"`) or by column/row coordinates
- **Table retrieval** -- extract data from named Excel tables
- **Discovery** -- list all sheet names and table names in a workbook

### RetrieveSheet Sample

```csharp
var smlDoc = new SmlDocument("data.xlsx");

var sheetNames = SmlDataRetriever.SheetNames(smlDoc);
Console.WriteLine($"Sheets: {string.Join(", ", sheetNames)}");

var sheetData = SmlDataRetriever.RetrieveSheet(smlDoc, sheetNames[0]);
Console.WriteLine(sheetData.ToString());
```

### RetrieveRange Sample

```csharp
var smlDoc = new SmlDocument("data.xlsx");

// By cell address range
var rangeData = SmlDataRetriever.RetrieveRange(smlDoc, "Sheet1", "A1:C10");

// By column/row coordinates (1-based)
var rangeData2 = SmlDataRetriever.RetrieveRange(
    smlDoc, "Sheet1",
    leftColumn: 1, topRow: 1, rightColumn: 3, bottomRow: 10);
```

### RetrieveTable Sample

```csharp
var smlDoc = new SmlDocument("data.xlsx");

var tableNames = SmlDataRetriever.TableNames(smlDoc);
foreach (var tableName in tableNames)
{
    var tableData = SmlDataRetriever.RetrieveTable(smlDoc, "Sheet1", tableName);
    Console.WriteLine($"Table: {tableName}");
    Console.WriteLine(tableData.ToString());
}
```
