// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Word;

public class ReferenceAdderTests : TestsBase
{
    // -----------------------------------------------------------------------
    // Helpers
    // -----------------------------------------------------------------------

    private static WmlDocument CreateMinimalDocumentWithHeadings()
    {
        using var stream = new MemoryStream();
        using (
            var wDoc = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
        )
        {
            var mainPart = wDoc.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    new Paragraph(
                        new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                        new Run(new Text("Chapter 1"))
                    ),
                    new Paragraph(new Run(new Text("Body text.")))
                )
            );

            // Minimal settings part (required so ReferenceAdder can write updateFields)
            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings();

            // Minimal styles part (required so ReferenceAdder can add TOC styles)
            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();

            // FontTablePart is required by ReferenceAdder.UpdateFontTablePart
            var fontTablePart = mainPart.AddNewPart<FontTablePart>();
            fontTablePart.Fonts = new DocumentFormat.OpenXml.Wordprocessing.Fonts();
        }
        return new WmlDocument("minimal.docx", stream.ToArray());
    }

    // -----------------------------------------------------------------------
    // AddToc — default title and tab position
    // -----------------------------------------------------------------------

    [Test]
    public async Task RA100_AddToc_DefaultTitle_InsertsContentControlBefore()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u", null, null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();

        var sdt = mainXDoc.Descendants(W.sdt).FirstOrDefault();
        await Assert.That(sdt).IsNotNull();

        // The SDT should appear before the heading paragraph
        var bodyChildren = mainXDoc.Root!.Element(W.body)!.Elements().ToList();
        var sdtIndex = bodyChildren.FindIndex(e => e.Name == W.sdt);
        var paraIndex = bodyChildren.FindIndex(e => e.Name == W.p);
        await Assert.That(sdtIndex).IsLessThan(paraIndex);
    }

    [Test]
    public async Task RA101_AddToc_DefaultTitle_IsContents()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u", null, null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();

        var headingPara = mainXDoc
            .Descendants(W.p)
            .FirstOrDefault(p => p.Descendants(W.pStyle).Any(ps => (string?)ps.Attribute(W.val) == "TOCHeading"));
        await Assert.That(headingPara).IsNotNull();

        var title = headingPara!.Descendants(W.t).Select(t => (string)t).FirstOrDefault();
        await Assert.That(title).IsEqualTo("Contents");
    }

    [Test]
    public async Task RA102_AddToc_CustomTitle_UsesProvidedTitle()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = ReferenceAdder.AddToc(
            doc,
            "/w:document/w:body/w:p[1]",
            @"TOC \o '1-3' \h \z \u",
            "Table of Contents",
            null
        );

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();

        var headingPara = mainXDoc
            .Descendants(W.p)
            .FirstOrDefault(p => p.Descendants(W.pStyle).Any(ps => (string?)ps.Attribute(W.val) == "TOCHeading"));
        await Assert.That(headingPara).IsNotNull();

        var title = headingPara!.Descendants(W.t).Select(t => (string)t).FirstOrDefault();
        await Assert.That(title).IsEqualTo("Table of Contents");
    }

    [Test]
    public async Task RA103_AddToc_InstrTextContainsSwitches()
    {
        const string switches = @"TOC \o '1-3' \h \z \u";
        var doc = CreateMinimalDocumentWithHeadings();
        var result = ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]", switches, null, null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();

        var instrText = mainXDoc.Descendants(W.instrText).Select(t => (string)t).FirstOrDefault();
        await Assert.That(instrText).IsNotNull();
        await Assert.That(instrText!.Trim()).IsEqualTo(switches);
    }

    [Test]
    public async Task RA104_AddToc_SetsUpdateFieldsInSettings()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u", null, null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var settingsXDoc = wDoc.MainDocumentPart!.DocumentSettingsPart!.GetXDocument();

        var updateFields = settingsXDoc.Descendants(W.updateFields).FirstOrDefault();
        await Assert.That(updateFields).IsNotNull();
        await Assert.That((string?)updateFields!.Attribute(W.val)).IsEqualTo("true");
    }

    [Test]
    public async Task RA105_AddToc_DocPartGalleryIsTableOfContents()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u", null, null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();

        var gallery = mainXDoc.Descendants(W.docPartGallery).FirstOrDefault();
        await Assert.That(gallery).IsNotNull();
        await Assert.That((string?)gallery!.Attribute(W.val)).IsEqualTo("Table of Contents");
    }

    [Test]
    public async Task RA106_AddToc_InvalidXPath_ThrowsException()
    {
        var doc = CreateMinimalDocumentWithHeadings();

        await Assert
            .That(() => ReferenceAdder.AddToc(doc, "/w:document/w:body/w:p[999]", @"TOC \o '1-3' \h \z \u", null, null))
            .Throws<OpenXmlPowerToolsException>();
    }

    // -----------------------------------------------------------------------
    // AddTof — table of figures
    // -----------------------------------------------------------------------

    [Test]
    public async Task RA110_AddTof_InsertsFieldBeforeTarget()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = ReferenceAdder.AddTof(doc, "/w:document/w:body/w:p[1]", @"TOC \h \z \c ""Figure""", null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();

        var instrTexts = mainXDoc.Descendants(W.instrText).Select(t => ((string)t).Trim()).ToList();
        await Assert.That(instrTexts).IsNotEmpty();
        await Assert.That(instrTexts[0]).Contains("Figure");

        var bodyChildren = mainXDoc.Root!.Element(W.body)!.Elements().ToList();
        var fieldParagraphIndex = bodyChildren.FindIndex(e =>
            e.Name == W.p && e.Descendants(W.instrText).Any(t => ((string)t).Contains("Figure"))
        );
        var targetParagraphIndex = bodyChildren.FindIndex(e =>
            e.Name == W.p && e.Descendants(W.t).Any(t => (string)t == "Chapter 1")
        );

        await Assert.That(fieldParagraphIndex).IsGreaterThanOrEqualTo(0);
        await Assert.That(targetParagraphIndex).IsGreaterThanOrEqualTo(0);
        await Assert.That(fieldParagraphIndex).IsLessThan(targetParagraphIndex);
    }

    [Test]
    public async Task RA111_AddTof_InvalidXPath_ThrowsException()
    {
        var doc = CreateMinimalDocumentWithHeadings();

        await Assert
            .That(() => ReferenceAdder.AddTof(doc, "/w:document/w:body/w:p[999]", @"TOC \h \z \c ""Figure""", null))
            .Throws<OpenXmlPowerToolsException>();
    }

    // -----------------------------------------------------------------------
    // AddToa — table of authorities
    // -----------------------------------------------------------------------

    [Test]
    public async Task RA120_AddToa_InsertsFieldBeforeTarget()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = ReferenceAdder.AddToa(doc, "/w:document/w:body/w:p[1]", @"TOA \h \c ""1"" \p", null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();

        var instrTexts = mainXDoc.Descendants(W.instrText).Select(t => ((string)t).Trim()).ToList();
        await Assert.That(instrTexts).IsNotEmpty();
        await Assert.That(instrTexts[0]).Contains("TOA");

        var bodyChildren = mainXDoc.Root!.Element(W.body)!.Elements().ToList();
        var fieldParagraphIndex = bodyChildren.FindIndex(e =>
            e.Name == W.p && e.Descendants(W.instrText).Any(t => ((string)t).Contains("TOA"))
        );
        var targetParagraphIndex = bodyChildren.FindIndex(e =>
            e.Name == W.p && e.Descendants(W.t).Any(t => (string)t == "Chapter 1")
        );

        await Assert.That(fieldParagraphIndex).IsGreaterThanOrEqualTo(0);
        await Assert.That(targetParagraphIndex).IsGreaterThanOrEqualTo(0);
        await Assert.That(fieldParagraphIndex).IsLessThan(targetParagraphIndex);
    }

    // -----------------------------------------------------------------------
    // WmlDocument extension methods
    // -----------------------------------------------------------------------

    [Test]
    public async Task RA130_WmlDocument_AddToc_ProducesValidDocument()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = doc.AddToc("/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u", "Contents", null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        await Validate(wDoc, []);
    }

    [Test]
    public async Task RA131_WmlDocument_AddTof_ProducesValidDocument()
    {
        var doc = CreateMinimalDocumentWithHeadings();
        var result = doc.AddTof("/w:document/w:body/w:p[1]", @"TOC \h \z \c ""Figure""", null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        await Validate(wDoc, []);
    }

    // -----------------------------------------------------------------------
    // File-based overloads (round-trip through existing test samples)
    // -----------------------------------------------------------------------

    [Test]
    [Arguments("RaTest01.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u")]
    [Arguments("RaTest02.docx", "/w:document/w:body/w:p[2]", @"TOC \o '1-3' \h \z \u")]
    public async Task RA140_AddToc_ExistingTestFiles_ProduceValidDocuments(
        string fileName,
        string xPath,
        string switches
    )
    {
        var srcFile = new FileInfo(Path.Combine("../../../Word/Samples/ReferenceAdder/", fileName));
        var srcDoc = new WmlDocument(srcFile.FullName);
        var result = ReferenceAdder.AddToc(srcDoc, xPath, switches, null, null);

        using var wDoc = WordprocessingDocument.Open(new MemoryStream(result.DocumentByteArray), false);
        await Validate(
            wDoc,
            [
                "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:updateFields'.",
            ]
        );

        var mainXDoc = wDoc.MainDocumentPart!.GetXDocument();
        var sdt = mainXDoc.Descendants(W.sdt).FirstOrDefault();
        await Assert.That(sdt).IsNotNull();
    }
}
