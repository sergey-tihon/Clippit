// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Excel;

public class SmlToHtmlConverterTests : TestsBase
{
    // PowerShell oneliner that generates InlineData for all files in a directory
    // dir | % { '[InlineData("' + $_.Name + '")]' } | clip
    [Test]
    [Arguments("SH101-SimpleFormats.xlsx", "Sheet1")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1")]
    [Arguments("SH103-No-SharedString.xlsx", "Sheet1")]
    [Arguments("SH104-With-SharedString.xlsx", "Sheet1")]
    [Arguments("SH105-No-SharedString.xlsx", "Sheet1")]
    [Arguments("SH106-9-x-9-Formatted.xlsx", "Sheet1")]
    [Arguments("SH108-SimpleFormattedCell.xlsx", "Sheet1")]
    [Arguments("SH109-CellWithBorder.xlsx", "Sheet1")]
    [Arguments("SH110-CellWithMasterStyle.xlsx", "Sheet1")]
    [Arguments("SH111-ChangedDefaultColumnWidth.xlsx", "Sheet1")]
    [Arguments("SH112-NotVertMergedCell.xlsx", "Sheet1")]
    [Arguments("SH113-VertMergedCell.xlsx", "Sheet1")]
    [Arguments("SH114-Centered-Cell.xlsx", "Sheet1")]
    [Arguments("SH115-DigitsToRight.xlsx", "Sheet1")]
    [Arguments("SH116-FmtNumId-1.xlsx", "Sheet1")]
    [Arguments("SH117-FmtNumId-2.xlsx", "Sheet1")]
    [Arguments("SH118-FmtNumId-3.xlsx", "Sheet1")]
    [Arguments("SH119-FmtNumId-4.xlsx", "Sheet1")]
    [Arguments("SH120-FmtNumId-9.xlsx", "Sheet1")]
    [Arguments("SH121-FmtNumId-11.xlsx", "Sheet1")]
    [Arguments("SH122-FmtNumId-12.xlsx", "Sheet1")]
    [Arguments("SH123-FmtNumId-14.xlsx", "Sheet1")]
    [Arguments("SH124-FmtNumId-15.xlsx", "Sheet1")]
    [Arguments("SH125-FmtNumId-16.xlsx", "Sheet1")]
    [Arguments("SH126-FmtNumId-17.xlsx", "Sheet1")]
    [Arguments("SH127-FmtNumId-18.xlsx", "Sheet1")]
    [Arguments("SH128-FmtNumId-19.xlsx", "Sheet1")]
    [Arguments("SH129-FmtNumId-20.xlsx", "Sheet1")]
    [Arguments("SH130-FmtNumId-21.xlsx", "Sheet1")]
    [Arguments("SH131-FmtNumId-22.xlsx", "Sheet1")]
    public void SH005_ConvertSheet(string name, string sheetName)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceXlsx = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var sourceCopiedToDestXlsx = new FileInfo(
            Path.Combine(TempDir, sourceXlsx.Name.Replace(".xlsx", "-1-Source.xlsx"))
        );
        if (!sourceCopiedToDestXlsx.Exists)
            File.Copy(sourceXlsx.FullName, sourceCopiedToDestXlsx.FullName);
        var dataTemplateFileNameSuffix = "-2-Generated-XmlData-Entire-Sheet.xml";
        var dataXmlFi = new FileInfo(
            Path.Combine(TempDir, sourceXlsx.Name.Replace(".xlsx", dataTemplateFileNameSuffix))
        );
        using var sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, false);
        var settings = new SmlToHtmlConverterSettings();
        var rangeXml = SmlDataRetriever.RetrieveSheet(sDoc, sheetName);
        rangeXml.Save(dataXmlFi.FullName);
    }

    [Test]
    [Arguments("SH101-SimpleFormats.xlsx", "Sheet1", "A1:B10")]
    [Arguments("SH101-SimpleFormats.xlsx", "Sheet1", "A4:B8")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1", "C2:C2")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1", "A9:A9")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1", "I1:I1")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1", "I9:I9")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1", "A1:I9")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1", "A2:D4")]
    [Arguments("SH102-9-x-9.xlsx", "Sheet1", "C5:G7")]
    [Arguments("SH103-No-SharedString.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH104-With-SharedString.xlsx", "Sheet1", "A4:A7")]
    [Arguments("SH105-No-SharedString.xlsx", "Sheet1", "A4:A7")]
    [Arguments("SH106-9-x-9-Formatted.xlsx", "Sheet1", "A1:I9")]
    [Arguments("SH108-SimpleFormattedCell.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH109-CellWithBorder.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH110-CellWithMasterStyle.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH111-ChangedDefaultColumnWidth.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH112-NotVertMergedCell.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH113-VertMergedCell.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH114-Centered-Cell.xlsx", "Sheet1", "A1:A1")]
    [Arguments("SH115-DigitsToRight.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH116-FmtNumId-1.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH117-FmtNumId-2.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH118-FmtNumId-3.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH119-FmtNumId-4.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH120-FmtNumId-9.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH121-FmtNumId-11.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH122-FmtNumId-12.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH123-FmtNumId-14.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH124-FmtNumId-15.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH125-FmtNumId-16.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH126-FmtNumId-17.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH127-FmtNumId-18.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH128-FmtNumId-19.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH129-FmtNumId-20.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH130-FmtNumId-21.xlsx", "Sheet1", "A1:A10")]
    [Arguments("SH131-FmtNumId-22.xlsx", "Sheet1", "A1:A10")]
    public Task SH004_ConvertRange(string name, string sheetName, string range)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceXlsx = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var sourceCopiedToDestXlsx = new FileInfo(
            Path.Combine(TempDir, sourceXlsx.Name.Replace(".xlsx", "-1-Source.xlsx"))
        );
        if (!sourceCopiedToDestXlsx.Exists)
            File.Copy(sourceXlsx.FullName, sourceCopiedToDestXlsx.FullName);
        var dataTemplateFileNameSuffix = $"-2-Generated-XmlData-{range.Replace(":", "")}.xml";
        var dataXmlFi = new FileInfo(
            Path.Combine(TempDir, sourceXlsx.Name.Replace(".xlsx", dataTemplateFileNameSuffix))
        );
        using var sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, false);
        var settings = new SmlToHtmlConverterSettings();
        var rangeXml = SmlDataRetriever.RetrieveRange(sDoc, sheetName, range);
        rangeXml.Save(dataXmlFi.FullName);
        return Task.CompletedTask;
    }

    [Test]
    [Arguments("SH001-Table.xlsx", "MyTable")]
    [Arguments("SH003-TableWithDateInFirstColumn.xlsx", "MyTable")]
    [Arguments("SH004-TableAtOffsetLocation.xlsx", "MyTable")]
    [Arguments("SH005-Table-With-SharedStrings.xlsx", "Table1")]
    [Arguments("SH006-Table-No-SharedStrings.xlsx", "Table1")]
    [Arguments("SH107-9-x-9-Formatted-Table.xlsx", "Table1")]
    [Arguments("SH007-One-Cell-Table.xlsx", "Table1")]
    [Arguments("SH008-Table-With-Tall-Row.xlsx", "Table1")]
    [Arguments("SH009-Table-With-Wide-Column.xlsx", "Table1")]
    public void SH003_ConvertTable(string name, string tableName)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceXlsx = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var sourceCopiedToDestXlsx = new FileInfo(
            Path.Combine(TempDir, sourceXlsx.Name.Replace(".xlsx", "-1-Source.xlsx"))
        );
        if (!sourceCopiedToDestXlsx.Exists)
            File.Copy(sourceXlsx.FullName, sourceCopiedToDestXlsx.FullName);
        var dataXmlFi = new FileInfo(
            Path.Combine(TempDir, sourceXlsx.Name.Replace(".xlsx", "-2-Generated-XmlData.xml"))
        );
        using var sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, false);
        //var settings = new SmlToHtmlConverterSettings();
        var rangeXml = SmlDataRetriever.RetrieveTable(sDoc, tableName);
        rangeXml.Save(dataXmlFi.FullName);
    }

    [Test]
    [Arguments("Spreadsheet.xlsx", 2)]
    public async Task SH002_SheetNames(string name, int numberOfSheets)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceXlsx = new FileInfo(Path.Combine(sourceDir.FullName, name));
        using var sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, false);
        var sheetNames = SmlDataRetriever.SheetNames(sDoc);
        await Assert.That(sheetNames).HasCount(numberOfSheets);
    }

    [Test]
    [Arguments("SH001-Table.xlsx", 1)]
    [Arguments("SH002-TwoTablesTwoSheets.xlsx", 2)]
    public async Task SH001_TableNames(string name, int numberOfTables)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceXlsx = new FileInfo(Path.Combine(sourceDir.FullName, name));
        using var sDoc = SpreadsheetDocument.Open(sourceXlsx.FullName, false);
        var table = SmlDataRetriever.TableNames(sDoc);
        await Assert.That(table).HasCount(numberOfTables);
    }
}
