// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Excel;

public class ExcelAssemblerTests : TestsBase
{
    // Helper: create a minimal template xlsx with cells containing placeholder strings.
    // Uses SpreadsheetWriter to build the base file with string cells (type "str").
    private static byte[] CreateTemplate(params (int row, int col, string value)[] cells)
    {
        var workbook = new WorkbookDfn
        {
            Worksheets =
            [
                new WorksheetDfn
                {
                    Name = "Sheet1",
                    Rows = cells
                        .GroupBy(c => c.row)
                        .OrderBy(g => g.Key)
                        .Select(g => new RowDfn { Cells = CreateRowCells(g) })
                        .ToArray(),
                },
            ],
        };
        using var ms = new MemoryStream();
        workbook.WriteTo(ms);
        return ms.ToArray();
    }

    private static CellDfn[] CreateRowCells(IGrouping<int, (int row, int col, string value)> rowCells)
    {
        var valuesByColumn = rowCells.ToDictionary(c => c.col, c => c.value);
        return Enumerable
            .Range(1, valuesByColumn.Keys.Max())
            .Select(col =>
                valuesByColumn.TryGetValue(col, out var value)
                    ? new CellDfn { CellDataType = CellDataType.String, Value = value }
                    : new CellDfn { CellDataType = CellDataType.String }
            )
            .ToArray();
    }

    // WorksheetAccessor.GetCellValue doesn't handle t="str" (formula-string cells written by
    // SpreadsheetWriter). This helper also covers that case.
    private static string? GetCellStringValue(SpreadsheetDocument doc, WorksheetPart ws, int column, int row)
    {
        var wsXDoc = ws.GetXDocument();
        var cellRef = WorksheetAccessor.GetColumnId(column) + row;
        var cell = wsXDoc.Descendants(S.c).FirstOrDefault(c => c.Attribute(NoNamespace.r)?.Value == cellRef);
        if (cell is null)
            return null;
        var t = cell.Attribute(NoNamespace.t)?.Value;
        return t switch
        {
            "s" => WorksheetAccessor.GetCellValue(doc, ws, column, row)?.ToString(),
            "inlineStr" => cell.Element(S._is)?.Element(S.t)?.Value,
            "str" => cell.Element(S.v)?.Value,
            _ => WorksheetAccessor.GetCellValue(doc, ws, column, row)?.ToString(),
        };
    }

    [Test]
    public async Task EA001_ScalarPlaceholderReplacement()
    {
        var templateBytes = CreateTemplate((1, 1, "{{Name}}"), (2, 1, "{{Age}}"));
        var data = XElement.Parse("<Root><Name>Alice</Name><Age>30</Age></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("Alice");
        await Assert.That(GetCellStringValue(doc, ws, 1, 2)).IsEqualTo("30");
    }

    [Test]
    public async Task EA002_MixedTextAndPlaceholder()
    {
        var templateBytes = CreateTemplate((1, 1, "Hello, {{Name}}!"));
        var data = XElement.Parse("<Root><Name>Bob</Name></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("Hello, Bob!");
    }

    [Test]
    public async Task EA003_MultiplePlaceholdersInOneCell()
    {
        var templateBytes = CreateTemplate((1, 1, "{{First}} {{Last}}"));
        var data = XElement.Parse("<Root><First>Jane</First><Last>Doe</Last></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("Jane Doe");
    }

    [Test]
    public async Task EA004_MissingXPathReturnsEmptyString()
    {
        var templateBytes = CreateTemplate((1, 1, "{{Missing}}"));
        var data = XElement.Parse("<Root><Name>Alice</Name></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task EA005_InvalidXPathProducesErrorMarker()
    {
        var templateBytes = CreateTemplate((1, 1, "{{[invalid}}"));
        var data = XElement.Parse("<Root/>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");
        var value = GetCellStringValue(doc, ws, 1, 1);

        await Assert.That(value).Contains("[XPathError:");
    }

    [Test]
    public async Task EA006_NonTemplateCellsAreUntouched()
    {
        var templateBytes = CreateTemplate((1, 1, "Static Label"), (1, 2, "{{Value}}"));
        var data = XElement.Parse("<Root><Value>42</Value></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        // Non-template cell stays as-is; template cell is replaced.
        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("Static Label");
        await Assert.That(GetCellStringValue(doc, ws, 2, 1)).IsEqualTo("42");
    }

    [Test]
    public async Task EA007_SmlDocumentOverload()
    {
        var templateBytes = CreateTemplate((1, 1, "{{Name}}"));
        var template = new SmlDocument("template.xlsx", templateBytes);
        var data = XElement.Parse("<Root><Name>Charlie</Name></Root>");

        var result = ExcelAssembler.AssembleDocument(template, data);

        await Assert.That(result.FileName).IsEqualTo("template.xlsx");
        using var doc = SpreadsheetDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");
        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("Charlie");
    }

    [Test]
    public async Task EA008_AttributeXPathResolution()
    {
        var templateBytes = CreateTemplate((1, 1, "{{Item/@id}}"));
        var data = XElement.Parse("<Root><Item id=\"123\">value</Item></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("123");
    }

    [Test]
    public async Task EA009_NonContiguousColumnIndexIsPreserved()
    {
        var templateBytes = CreateTemplate((1, 3, "{{Name}}"));
        var data = XElement.Parse("<Root><Name>Alice</Name></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        await Assert.That(GetCellStringValue(doc, ws, 3, 1)).IsEqualTo("Alice");
    }
}
