// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Xml.Linq;
using Clippit.Excel;
using Clippit.PowerPoint;
using Clippit.Word;
using TUnit.Assertions;
using TUnit.Assertions.Extensions;
using TUnit.Core;

namespace Clippit.Tests.Common;

public class MetricsGetterTests : TestsBase
{
    private static readonly DirectoryInfo s_sourceDir = new("../../../../TestFiles/");

    [Test]
    [Arguments("Presentation.pptx")]
    [Arguments("Spreadsheet.xlsx")]
    [Arguments("DA/DA001-TemplateDocument.docx")]
    [Arguments("DA/DA002-TemplateDocument.docx")]
    [Arguments("DA/DA003-Select-XPathFindsNoData.docx")]
    [Arguments("DA/DA004-Select-XPathFindsNoDataOptional.docx")]
    [Arguments("DA/DA005-SelectRowData-NoData.docx")]
    [Arguments("DA/DA006-SelectTestValue-NoData.docx")]
    public async Task MG001(string name)
    {
        var fi = new FileInfo(Path.Combine(s_sourceDir.FullName, name));
        var settings = new MetricsGetterSettings()
        {
            IncludeTextInContentControls = false,
            IncludeXlsxTableCellData = false,
            RetrieveNamespaceList = true,
            RetrieveContentTypeList = true,
        };
        var extension = fi.Extension.ToLower();
        XElement metrics = null;
        if (Util.IsWordprocessingML(extension))
        {
            var wmlDocument = new WmlDocument(fi.FullName);
            metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);
        }
        else if (Util.IsSpreadsheetML(extension))
        {
            var smlDocument = new SmlDocument(fi.FullName);
            metrics = MetricsGetter.GetXlsxMetrics(smlDocument, settings);
        }
        else if (Util.IsPresentationML(extension))
        {
            var pmlDocument = new PmlDocument(fi.FullName);
            metrics = MetricsGetter.GetPptxMetrics(pmlDocument, settings);
        }

        await Assert.That(metrics).IsNotNull();
    }

    // ── GetDocxMetrics: structure ────────────────────────────────────────────

    [Test]
    public async Task MG002_GetDocxMetrics_RootElement_IsMetrics()
    {
        var wmlDoc = new WmlDocument(Path.Combine(s_sourceDir.FullName, "DA/DA001-TemplateDocument.docx"));
        var metrics = MetricsGetter.GetDocxMetrics(wmlDoc, new MetricsGetterSettings());
        await Assert.That(metrics.Name).IsEqualTo((XName)"Metrics");
    }

    [Test]
    public async Task MG003_GetDocxMetrics_FileTypeAttribute_IsWordprocessingML()
    {
        var wmlDoc = new WmlDocument(Path.Combine(s_sourceDir.FullName, "DA/DA001-TemplateDocument.docx"));
        var metrics = MetricsGetter.GetDocxMetrics(wmlDoc, new MetricsGetterSettings());
        await Assert.That((string)metrics.Attribute("FileType")).IsEqualTo("WordprocessingML");
    }

    [Test]
    public async Task MG004_GetDocxMetrics_FileNameAttribute_ContainsFileName()
    {
        var wmlDoc = new WmlDocument(Path.Combine(s_sourceDir.FullName, "DA/DA001-TemplateDocument.docx"));
        var metrics = MetricsGetter.GetDocxMetrics(wmlDoc, new MetricsGetterSettings());
        await Assert.That((string)metrics.Attribute("FileName")).Contains("DA001-TemplateDocument.docx");
    }

    [Test]
    public async Task MG005_GetDocxMetrics_RetrieveContentTypeList_AddsContentTypesElement()
    {
        var wmlDoc = new WmlDocument(Path.Combine(s_sourceDir.FullName, "DA/DA001-TemplateDocument.docx"));
        var settings = new MetricsGetterSettings { RetrieveContentTypeList = true };
        var metrics = MetricsGetter.GetDocxMetrics(wmlDoc, settings);
        await Assert.That(metrics.Element("ContentTypes")).IsNotNull();
        await Assert.That(metrics.Element("ContentTypes")!.HasElements).IsTrue();
    }

    [Test]
    public async Task MG006_GetDocxMetrics_RetrieveNamespaceList_AddsNamespacesElement()
    {
        var wmlDoc = new WmlDocument(Path.Combine(s_sourceDir.FullName, "DA/DA001-TemplateDocument.docx"));
        var settings = new MetricsGetterSettings { RetrieveNamespaceList = true };
        var metrics = MetricsGetter.GetDocxMetrics(wmlDoc, settings);
        await Assert.That(metrics.Element("Namespaces")).IsNotNull();
        await Assert.That(metrics.Element("Namespaces")!.HasElements).IsTrue();
    }

    [Test]
    public async Task MG007_GetDocxMetrics_NoContentTypesWhenDisabled()
    {
        var wmlDoc = new WmlDocument(Path.Combine(s_sourceDir.FullName, "DA/DA001-TemplateDocument.docx"));
        var settings = new MetricsGetterSettings { RetrieveContentTypeList = false };
        var metrics = MetricsGetter.GetDocxMetrics(wmlDoc, settings);
        await Assert.That(metrics.Element("ContentTypes")).IsNull();
    }

    [Test]
    public async Task MG008_GetDocxMetrics_NoNamespacesWhenDisabled()
    {
        var wmlDoc = new WmlDocument(Path.Combine(s_sourceDir.FullName, "DA/DA001-TemplateDocument.docx"));
        var settings = new MetricsGetterSettings { RetrieveNamespaceList = false };
        var metrics = MetricsGetter.GetDocxMetrics(wmlDoc, settings);
        await Assert.That(metrics.Element("Namespaces")).IsNull();
    }

    // ── GetXlsxMetrics: structure ────────────────────────────────────────────

    [Test]
    public async Task MG009_GetXlsxMetrics_RootElement_IsMetrics()
    {
        var smlDoc = new SmlDocument(Path.Combine(s_sourceDir.FullName, "Spreadsheet.xlsx"));
        var metrics = MetricsGetter.GetXlsxMetrics(smlDoc, new MetricsGetterSettings());
        await Assert.That(metrics.Name).IsEqualTo((XName)"Metrics");
    }

    [Test]
    public async Task MG010_GetXlsxMetrics_FileTypeAttribute_IsSpreadsheetML()
    {
        var smlDoc = new SmlDocument(Path.Combine(s_sourceDir.FullName, "Spreadsheet.xlsx"));
        var metrics = MetricsGetter.GetXlsxMetrics(smlDoc, new MetricsGetterSettings());
        await Assert.That((string)metrics.Attribute("FileType")).IsEqualTo("SpreadsheetML");
    }

    // ── GetPptxMetrics: structure ────────────────────────────────────────────

    [Test]
    public async Task MG011_GetPptxMetrics_RootElement_IsMetrics()
    {
        var pmlDoc = new PmlDocument(Path.Combine(s_sourceDir.FullName, "Presentation.pptx"));
        var metrics = MetricsGetter.GetPptxMetrics(pmlDoc, new MetricsGetterSettings());
        await Assert.That(metrics.Name).IsEqualTo((XName)"Metrics");
    }

    [Test]
    public async Task MG012_GetPptxMetrics_FileTypeAttribute_IsPresentationML()
    {
        var pmlDoc = new PmlDocument(Path.Combine(s_sourceDir.FullName, "Presentation.pptx"));
        var metrics = MetricsGetter.GetPptxMetrics(pmlDoc, new MetricsGetterSettings());
        await Assert.That((string)metrics.Attribute("FileType")).IsEqualTo("PresentationML");
    }

    // ── GetMetrics (unified): auto-dispatch by extension ────────────────────

    [Test]
    [Arguments("DA/DA001-TemplateDocument.docx", "WordprocessingML")]
    [Arguments("Spreadsheet.xlsx", "SpreadsheetML")]
    [Arguments("Presentation.pptx", "PresentationML")]
    public async Task MG013_GetMetrics_ByExtension_ReturnsCorrectFileType(string name, string expectedFileType)
    {
        var path = Path.Combine(s_sourceDir.FullName, name);
        var metrics = MetricsGetter.GetMetrics(path, new MetricsGetterSettings());
        await Assert.That(metrics).IsNotNull();
        await Assert.That((string)metrics!.Attribute("FileType")).IsEqualTo(expectedFileType);
    }

    [Test]
    public async Task MG014_GetMetrics_UnknownExtension_ReturnsNull()
    {
        // txt files are not Office documents — GetMetrics returns null
        var tempTxt = Path.Combine(TempDir, "dummy.txt");
        File.WriteAllText(tempTxt, "hello");
        var metrics = MetricsGetter.GetMetrics(tempTxt, new MetricsGetterSettings());
        await Assert.That(metrics).IsNull();
    }
}
