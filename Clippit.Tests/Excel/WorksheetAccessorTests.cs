// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Excel;

/// <summary>
/// Unit tests for <see cref="WorksheetAccessor.GetColumnId"/> and
/// <see cref="WorksheetAccessor.GetColumnNumber"/> — the 1-based column-address
/// utility methods used throughout spreadsheet manipulation.
/// </summary>
public class WorksheetAccessorTests
{
    // ── GetColumnId ──────────────────────────────────────────────────────────
    // Converts a 1-based column number to its Excel column-letter identifier.

    [Test]
    [Arguments(1, "A")]
    [Arguments(2, "B")]
    [Arguments(26, "Z")]
    [Arguments(27, "AA")]
    [Arguments(28, "AB")]
    [Arguments(52, "AZ")]
    [Arguments(53, "BA")]
    [Arguments(702, "ZZ")]
    [Arguments(703, "AAA")]
    [Arguments(704, "AAB")]
    [Arguments(18278, "ZZZ")]
    public async Task WA001_GetColumnId_KnownValues_ReturnsExpectedLetter(int column, string expected)
    {
        var result = WorksheetAccessor.GetColumnId(column);
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── GetColumnNumber ──────────────────────────────────────────────────────
    // Extracts the 1-based column number from a cell reference such as "A1" or "AA99".
    // Digit characters (the row number) are ignored; only letters are used.

    [Test]
    [Arguments("A1", 1)]
    [Arguments("B1", 2)]
    [Arguments("Z1", 26)]
    [Arguments("AA1", 27)]
    [Arguments("AB1", 28)]
    [Arguments("AZ1", 52)]
    [Arguments("BA1", 53)]
    [Arguments("ZZ999", 702)]
    [Arguments("AAA1", 703)]
    [Arguments("ZZZ1048576", 18278)]
    public async Task WA002_GetColumnNumber_CellReference_ReturnsExpected1BasedColumn(
        string cellReference,
        int expected
    )
    {
        var result = WorksheetAccessor.GetColumnNumber(cellReference);
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── Round-trip: GetColumnId → GetColumnNumber ────────────────────────────

    [Test]
    [Arguments(1)]
    [Arguments(26)]
    [Arguments(27)]
    [Arguments(52)]
    [Arguments(702)]
    [Arguments(703)]
    [Arguments(18278)]
    public async Task WA003_RoundTrip_ColumnId_ThenColumnNumber_IsIdentity(int originalColumn)
    {
        var address = WorksheetAccessor.GetColumnId(originalColumn) + "1";
        var recovered = WorksheetAccessor.GetColumnNumber(address);
        await Assert.That(recovered).IsEqualTo(originalColumn);
    }

    // ── SetCellValue / GetCellValue round-trip ───────────────────────────────

    private static SpreadsheetDocument CreateBlankSpreadsheet(MemoryStream ms)
    {
        var doc = SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
        doc.AddWorkbookPart();
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        doc.WorkbookPart!.PutXDocument(
            new XDocument(
                new XElement(
                    ns + "workbook",
                    new XAttribute("xmlns", ns),
                    new XAttribute(XNamespace.Xmlns + "r", r),
                    new XElement(ns + "sheets")
                )
            )
        );
        return doc;
    }

    [Test]
    public async Task WA004_SetGetCellValue_Int_RoundTrip()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 1, value: 42);
        sheet.PutXDocument();

        var result = WorksheetAccessor.GetCellValue(doc, sheet, column: 1, row: 1);

        await Assert.That(result).IsEqualTo(42);
    }

    [Test]
    public async Task WA005_SetGetCellValue_Double_RoundTrip()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        WorksheetAccessor.SetCellValue(doc, sheet, row: 2, column: 3, value: 3.14);
        sheet.PutXDocument();

        var result = WorksheetAccessor.GetCellValue(doc, sheet, column: 3, row: 2);

        await Assert.That(result).IsEqualTo(3.14);
    }

    [Test]
    public async Task WA006_SetGetCellValue_String_RoundTrip()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 2, value: "Hello");
        sheet.PutXDocument();

        var result = WorksheetAccessor.GetCellValue(doc, sheet, column: 2, row: 1);

        await Assert.That(result).IsEqualTo("Hello");
    }

    [Test]
    public async Task WA007_SetGetCellValue_Bool_True_RoundTrip()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 1, value: true);
        sheet.PutXDocument();

        var result = WorksheetAccessor.GetCellValue(doc, sheet, column: 1, row: 1);

        await Assert.That(result).IsEqualTo(true);
    }

    [Test]
    public async Task WA008_SetGetCellValue_Bool_False_RoundTrip()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 1, value: false);
        sheet.PutXDocument();

        var result = WorksheetAccessor.GetCellValue(doc, sheet, column: 1, row: 1);

        await Assert.That(result).IsEqualTo(false);
    }

    [Test]
    public async Task WA009_GetCellValue_NonExistentCell_ReturnsNull()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        var result = WorksheetAccessor.GetCellValue(doc, sheet, column: 5, row: 99);

        await Assert.That(result).IsNull();
    }

    [Test]
    public async Task WA010_SetCellValue_OverwriteExistingCell_UpdatesValue()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 1, value: "Original");
        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 1, value: "Updated");
        sheet.PutXDocument();

        var result = WorksheetAccessor.GetCellValue(doc, sheet, column: 1, row: 1);

        await Assert.That(result).IsEqualTo("Updated");
    }

    [Test]
    public async Task WA011_SetCellValue_MultipleCellsSameRow_AllRetrievableIndependently()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 1, value: "Name");
        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 2, value: 100);
        WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 3, value: true);
        sheet.PutXDocument();

        await Assert.That(WorksheetAccessor.GetCellValue(doc, sheet, column: 1, row: 1)).IsEqualTo("Name");
        await Assert.That(WorksheetAccessor.GetCellValue(doc, sheet, column: 2, row: 1)).IsEqualTo(100);
        await Assert.That(WorksheetAccessor.GetCellValue(doc, sheet, column: 3, row: 1)).IsEqualTo(true);
    }

    [Test]
    public async Task WA012_AddWorksheet_GetWorksheet_RoundTrip()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);

        _ = WorksheetAccessor.AddWorksheet(doc, "MySheet");

        var retrieved = WorksheetAccessor.GetWorksheet(doc, "MySheet");

        await Assert.That(retrieved).IsNotNull();
    }

    [Test]
    public async Task WA013_SetCellValue_InvalidType_ThrowsArgumentException()
    {
        using var ms = new MemoryStream();
        using var doc = CreateBlankSpreadsheet(ms);
        var sheet = WorksheetAccessor.AddWorksheet(doc, "Data");

        await Assert
            .That(() => WorksheetAccessor.SetCellValue(doc, sheet, row: 1, column: 1, value: new object()))
            .Throws<ArgumentException>();
    }
}
