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
        if (valuesByColumn.Count == 0)
            return [];

        return Enumerable
            .Range(1, valuesByColumn.Keys.Max())
            .Select(col =>
                valuesByColumn.TryGetValue(col, out var value)
                    ? new CellDfn { CellDataType = CellDataType.String, Value = value }
                    : new CellDfn { CellDataType = CellDataType.String }
            )
            .ToArray();
    }

    private static byte[] CreateSharedStringTemplate(string value)
    {
        var templateBytes = CreateTemplate((1, 1, value));
        using var ms = new MemoryStream();
        ms.Write(templateBytes, 0, templateBytes.Length);
        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, true))
        {
            var workbookPart = doc.WorkbookPart;
            if (workbookPart is null)
                throw new InvalidOperationException("WorkbookPart is missing in test template.");

            var wsPart = workbookPart.WorksheetParts.First();
            var wsXDoc = wsPart.GetXDocument();
            var cell = wsXDoc.Descendants(S.c).First();
            cell.SetAttributeValue(NoNamespace.t, "s");
            cell.Elements(S.v).Remove();
            cell.Elements(S._is).Remove();
            cell.Add(new XElement(S.v, "0"));
            wsPart.PutXDocument();

            var sstPart = workbookPart.SharedStringTablePart ?? workbookPart.AddNewPart<SharedStringTablePart>();
            sstPart.PutXDocument(
                new XDocument(
                    new XElement(
                        S.sst,
                        new XAttribute(NoNamespace.count, "1"),
                        new XAttribute(NoNamespace.uniqueCount, "1"),
                        new XElement(S.si, new XElement(S.t, value))
                    )
                )
            );
        }
        return ms.ToArray();
    }

    private static byte[] CreateInlineRichStringTemplate(params string[] runTexts)
    {
        var templateBytes = CreateTemplate((1, 1, "placeholder"));
        using var ms = new MemoryStream();
        ms.Write(templateBytes, 0, templateBytes.Length);
        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, true))
        {
            var wsPart =
                doc.WorkbookPart?.WorksheetParts.First()
                ?? throw new InvalidOperationException("WorksheetPart is missing in test template.");
            var wsXDoc = wsPart.GetXDocument();
            var cell = wsXDoc.Descendants(S.c).First();
            cell.SetAttributeValue(NoNamespace.t, "inlineStr");
            cell.Elements(S.v).Remove();
            cell.Elements(S._is).Remove();
            cell.Add(new XElement(S._is, runTexts.Select(text => new XElement(S.r, new XElement(S.t, text)))));
            wsPart.PutXDocument();
        }
        return ms.ToArray();
    }

    private static byte[] CreateFormulaStringTemplate(string formulaResult)
    {
        var templateBytes = CreateTemplate((1, 1, formulaResult));
        using var ms = new MemoryStream();
        ms.Write(templateBytes, 0, templateBytes.Length);
        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, true))
        {
            var wsPart =
                doc.WorkbookPart?.WorksheetParts.First()
                ?? throw new InvalidOperationException("WorksheetPart is missing in test template.");
            var wsXDoc = wsPart.GetXDocument();
            var cell = wsXDoc.Descendants(S.c).First();
            cell.SetAttributeValue(NoNamespace.t, "str");
            cell.Elements(S.f).Remove();
            cell.AddFirst(new XElement(S.f, "\"seed\""));
            wsPart.PutXDocument();
        }
        return ms.ToArray();
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

    [Test]
    public async Task EA010_SharedStringCellTemplateReplacement()
    {
        var templateBytes = CreateSharedStringTemplate("{{Name}}");
        var data = XElement.Parse("<Root><Name>Dana</Name></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("Dana");
    }

    [Test]
    public async Task EA011_InlineRichStringTemplateReplacement()
    {
        var templateBytes = CreateInlineRichStringTemplate("Hello, ", "{{Name}}", "!");
        var data = XElement.Parse("<Root><Name>Eve</Name></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");

        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("Hello, Eve!");
    }

    [Test]
    public async Task EA012_TemplateFormulaIsRemovedWhenCellIsRewritten()
    {
        var templateBytes = CreateFormulaStringTemplate("{{Name}}");
        var data = XElement.Parse("<Root><Name>Frank</Name></Root>");

        var resultBytes = ExcelAssembler.AssembleDocument(templateBytes, data);

        using var doc = SpreadsheetDocument.Open(new MemoryStream(resultBytes), false);
        var ws = WorksheetAccessor.GetWorksheet(doc, "Sheet1");
        var wsXDoc = ws.GetXDocument();
        var cell = wsXDoc.Descendants(S.c).First(c => c.Attribute(NoNamespace.r)?.Value == "A1");

        await Assert.That(GetCellStringValue(doc, ws, 1, 1)).IsEqualTo("Frank");
        await Assert.That(cell.Element(S.f)).IsNull();
    }
}
