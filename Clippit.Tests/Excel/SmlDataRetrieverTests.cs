// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Excel;

/// <summary>
/// Unit tests for <see cref="SmlDataRetriever"/>:
/// <list type="bullet">
/// <item><description><see cref="SmlDataRetriever.SheetNames(SmlDocument)"/> — enumerates worksheet names</description></item>
/// <item><description><see cref="SmlDataRetriever.TableNames(SmlDocument)"/> — enumerates table names</description></item>
/// <item><description><see cref="SmlDataRetriever.RetrieveSheet(SmlDocument,string)"/> — retrieves all cells in a sheet as XML</description></item>
/// <item><description><see cref="SmlDataRetriever.RetrieveRange(SmlDocument,string,string)"/> — retrieves a rectangular range as XML</description></item>
/// <item><description><see cref="SmlDataRetriever.RetrieveTable(SmlDocument,string,string)"/> — retrieves a named table as XML</description></item>
/// </list>
/// </summary>
public class SmlDataRetrieverTests : TestsBase
{
    private static readonly DirectoryInfo SourceDir = new("../../../../TestFiles/");

    private static string Path(string name) => System.IO.Path.Combine(SourceDir.FullName, name);

    // ── SheetNames ───────────────────────────────────────────────────────────

    [Test]
    public async Task SDR001_SheetNames_SingleSheet_ReturnsOneEntry()
    {
        var names = SmlDataRetriever.SheetNames(Path("SH001-Table.xlsx"));
        await Assert.That(names).HasCount().EqualTo(1);
        await Assert.That(names[0]).IsEqualTo("Sheet1");
    }

    [Test]
    public async Task SDR002_SheetNames_TwoSheets_ReturnsBothInOrder()
    {
        var names = SmlDataRetriever.SheetNames(Path("SH002-TwoTablesTwoSheets.xlsx"));
        await Assert.That(names).HasCount().EqualTo(2);
        await Assert.That(names).Contains("Sheet1");
        await Assert.That(names).Contains("Sheet2");
        await Assert.That(names[0]).IsEqualTo("Sheet1");
    }

    [Test]
    public async Task SDR003_SheetNames_SmlDocumentOverload_MatchesFileOverload()
    {
        var filePath = Path("SH002-TwoTablesTwoSheets.xlsx");
        var smlDoc = new SmlDocument(filePath);
        var fromFile = SmlDataRetriever.SheetNames(filePath);
        var fromSml = SmlDataRetriever.SheetNames(smlDoc);
        await Assert.That(fromSml).IsEquivalentTo(fromFile);
    }

    // ── TableNames ───────────────────────────────────────────────────────────

    [Test]
    public async Task SDR004_TableNames_SingleTable_ReturnsOneName()
    {
        var names = SmlDataRetriever.TableNames(Path("SH001-Table.xlsx"));
        await Assert.That(names).HasCount().EqualTo(1);
        await Assert.That(names[0]).IsEqualTo("MyTable");
    }

    [Test]
    public async Task SDR005_TableNames_TwoTables_ReturnsBothNames()
    {
        var names = SmlDataRetriever.TableNames(Path("SH002-TwoTablesTwoSheets.xlsx"));
        await Assert.That(names).HasCount().EqualTo(2);
        await Assert.That(names).Contains("MyTable");
        await Assert.That(names).Contains("MyTable2");
    }

    [Test]
    public async Task SDR006_TableNames_SmlDocumentOverload_MatchesFileOverload()
    {
        var filePath = Path("SH001-Table.xlsx");
        var smlDoc = new SmlDocument(filePath);
        var fromFile = SmlDataRetriever.TableNames(filePath);
        var fromSml = SmlDataRetriever.TableNames(smlDoc);
        await Assert.That(fromSml).IsEquivalentTo(fromFile);
    }

    // ── RetrieveSheet ────────────────────────────────────────────────────────

    [Test]
    public async Task SDR007_RetrieveSheet_ReturnsDataElement()
    {
        var result = SmlDataRetriever.RetrieveSheet(Path("SH001-Table.xlsx"), "Sheet1");
        await Assert.That(result).IsNotNull();
        await Assert.That(result.Name.LocalName).IsEqualTo("Data");
    }

    [Test]
    public async Task SDR008_RetrieveSheet_ReturnsExpectedRowCount()
    {
        // SH001-Table.xlsx has range A1:C3 (header + 2 data rows = 3 rows total)
        var result = SmlDataRetriever.RetrieveSheet(Path("SH001-Table.xlsx"), "Sheet1");
        var rows = result.Elements("Row").ToList();
        await Assert.That(rows).HasCount().EqualTo(3);
    }

    [Test]
    public async Task SDR009_RetrieveSheet_RowsHaveRowNumberAttribute()
    {
        var result = SmlDataRetriever.RetrieveSheet(Path("SH001-Table.xlsx"), "Sheet1");
        foreach (var row in result.Elements("Row"))
        {
            await Assert.That(row.Attribute("RowNumber")).IsNotNull();
        }
    }

    [Test]
    public async Task SDR010_RetrieveSheet_CellsHaveValueElements()
    {
        var result = SmlDataRetriever.RetrieveSheet(Path("SH001-Table.xlsx"), "Sheet1");
        var cells = result.Descendants("Cell").ToList();
        await Assert.That(cells).IsNotEmpty();
        foreach (var cell in cells)
        {
            await Assert.That(cell.Element("Value")).IsNotNull();
            await Assert.That(cell.Element("DisplayValue")).IsNotNull();
        }
    }

    [Test]
    public async Task SDR011_RetrieveSheet_InvalidSheetName_ThrowsArgumentException()
    {
        await Assert
            .That(() => SmlDataRetriever.RetrieveSheet(Path("SH001-Table.xlsx"), "DoesNotExist"))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task SDR012_RetrieveSheet_SmlDocumentOverload_ReturnsData()
    {
        var smlDoc = new SmlDocument(Path("SH001-Table.xlsx"));
        var result = SmlDataRetriever.RetrieveSheet(smlDoc, "Sheet1");
        await Assert.That(result.Name.LocalName).IsEqualTo("Data");
        await Assert.That(result.Elements("Row")).IsNotEmpty();
    }

    // ── RetrieveRange ────────────────────────────────────────────────────────

    [Test]
    public async Task SDR013_RetrieveRange_SingleCell_ReturnsOneRow()
    {
        var result = SmlDataRetriever.RetrieveRange(Path("SH001-Table.xlsx"), "Sheet1", "A1:A1");
        var rows = result.Elements("Row").ToList();
        await Assert.That(rows).HasCount().EqualTo(1);
        var cells = rows[0].Elements("Cell").ToList();
        await Assert.That(cells).HasCount().EqualTo(1);
    }

    [Test]
    public async Task SDR014_RetrieveRange_OneColumn_ReturnsAllRowsInRange()
    {
        // A1:A3 = 3 rows, each with 1 cell
        var result = SmlDataRetriever.RetrieveRange(Path("SH001-Table.xlsx"), "Sheet1", "A1:A3");
        var rows = result.Elements("Row").ToList();
        await Assert.That(rows).HasCount().EqualTo(3);
        foreach (var row in rows)
        {
            await Assert.That(row.Elements("Cell").ToList()).HasCount().EqualTo(1);
        }
    }

    [Test]
    public async Task SDR015_RetrieveRange_InvalidSheetName_ThrowsArgumentException()
    {
        await Assert
            .That(() => SmlDataRetriever.RetrieveRange(Path("SH001-Table.xlsx"), "BadSheet", "A1:C3"))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task SDR016_RetrieveRange_SmlDocumentOverload_MatchesFileOverload()
    {
        var filePath = Path("SH001-Table.xlsx");
        var smlDoc = new SmlDocument(filePath);
        var fromFile = SmlDataRetriever.RetrieveRange(filePath, "Sheet1", "A1:C2");
        var fromSml = SmlDataRetriever.RetrieveRange(smlDoc, "Sheet1", "A1:C2");
        await Assert.That(XNode.DeepEquals(fromSml, fromFile)).IsTrue();
    }

    // ── RetrieveTable ────────────────────────────────────────────────────────

    [Test]
    public async Task SDR017_RetrieveTable_ReturnsTableElement()
    {
        var result = SmlDataRetriever.RetrieveTable(Path("SH001-Table.xlsx"), "MyTable");
        await Assert.That(result).IsNotNull();
        await Assert.That(result.Name.LocalName).IsEqualTo("Table");
    }

    [Test]
    public async Task SDR018_RetrieveTable_HasTableNameAttribute()
    {
        var result = SmlDataRetriever.RetrieveTable(Path("SH001-Table.xlsx"), "MyTable");
        await Assert.That((string)result.Attribute("TableName")).IsEqualTo("MyTable");
    }

    [Test]
    public async Task SDR019_RetrieveTable_HasColumnsElement()
    {
        var result = SmlDataRetriever.RetrieveTable(Path("SH001-Table.xlsx"), "MyTable");
        var columns = result.Element("Columns");
        await Assert.That(columns).IsNotNull();
        await Assert.That(columns!.Elements("Column").ToList()).IsNotEmpty();
    }

    [Test]
    public async Task SDR020_RetrieveTable_DataRowCountExcludesHeader()
    {
        // SH001-Table.xlsx: table A1:C3, 1 header row → 2 data rows
        var result = SmlDataRetriever.RetrieveTable(Path("SH001-Table.xlsx"), "MyTable");
        var dataRows = result.Element("Data")?.Elements("Row").ToList();
        await Assert.That(dataRows).IsNotNull();
        await Assert.That(dataRows!).HasCount().EqualTo(2);
    }

    [Test]
    public async Task SDR021_RetrieveTable_InvalidTableName_ThrowsArgumentException()
    {
        await Assert
            .That(() => SmlDataRetriever.RetrieveTable(Path("SH001-Table.xlsx"), "NoSuchTable"))
            .Throws<ArgumentException>();
    }

    [Test]
    public async Task SDR022_RetrieveTable_SmlDocumentOverload_ReturnsTable()
    {
        var smlDoc = new SmlDocument(Path("SH001-Table.xlsx"));
        var result = SmlDataRetriever.RetrieveTable(smlDoc, "Sheet1", "MyTable");
        await Assert.That(result.Name.LocalName).IsEqualTo("Table");
        await Assert.That((string)result.Attribute("TableName")).IsEqualTo("MyTable");
    }

    [Test]
    public async Task SDR023_RetrieveTable_SharedStrings_ValuesAreDecoded()
    {
        // SH005-Table-With-SharedStrings.xlsx has shared strings
        var result = SmlDataRetriever.RetrieveTable(Path("SH005-Table-With-SharedStrings.xlsx"), "Table1");
        var cells = result.Descendants("Cell").ToList();
        await Assert.That(cells).IsNotEmpty();
        // Every Cell element must have a non-null Value child
        foreach (var cell in cells)
        {
            await Assert.That(cell.Element("Value")).IsNotNull();
        }
        var sharedStringValues = cells
            .Where(cell => (string?)cell.Attribute("Type") == "s")
            .Select(cell => (string?)cell.Element("Value"))
            .Where(value => !string.IsNullOrEmpty(value))
            .Cast<string>()
            .ToList();
        await Assert.That(sharedStringValues).IsNotEmpty();
        await Assert.That(sharedStringValues.Any(value => !int.TryParse(value, out _))).IsTrue();
    }

    [Test]
    public async Task SDR024_SpreadsheetDocument_Overload_SheetNames()
    {
        using var sDoc = SpreadsheetDocument.Open(Path("SH002-TwoTablesTwoSheets.xlsx"), false);
        var sheetNames = SmlDataRetriever.SheetNames(sDoc);
        await Assert.That(sheetNames).HasCount().EqualTo(2);
        await Assert.That(sheetNames).Contains("Sheet1");
        await Assert.That(sheetNames).Contains("Sheet2");
        var tableNames = SmlDataRetriever.TableNames(sDoc);
        await Assert.That(tableNames).HasCount().EqualTo(2);
        await Assert.That(tableNames).Contains("MyTable");
        await Assert.That(tableNames).Contains("MyTable2");
    }

    [Test]
    public async Task SDR025_SpreadsheetDocument_Overload_RetrieveSheet()
    {
        using var sDoc = SpreadsheetDocument.Open(Path("SH001-Table.xlsx"), false);
        var sheet = SmlDataRetriever.RetrieveSheet(sDoc, "Sheet1");
        await Assert.That(sheet.Name.LocalName).IsEqualTo("Data");
        await Assert.That(sheet.Elements("Row").ToList()).HasCount().EqualTo(3);
        var range = SmlDataRetriever.RetrieveRange(sDoc, "Sheet1", "A1:C2");
        var rows = range.Elements("Row").ToList();
        await Assert.That(rows).HasCount().EqualTo(2);
        await Assert.That(rows.All(row => row.Elements("Cell").Count() == 3)).IsTrue();
        var table = SmlDataRetriever.RetrieveTable(sDoc, "MyTable");
        await Assert.That(table.Name.LocalName).IsEqualTo("Table");
        await Assert.That((string?)table.Attribute("TableName")).IsEqualTo("MyTable");
        await Assert.That(table.Element("Data")?.Elements("Row").ToList()).HasCount().EqualTo(2);
    }
}
