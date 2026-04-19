// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
#define COPY_FILES_FOR_DEBUGGING
// DO_CONVERSION_VIA_WORD is defined in the project Clippit.Tests.OA.csproj, but not in the Clippit.Tests.csproj
using System.Text;
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Word;

public class HtmlConverterTests() : Clippit.Tests.TestsBase
{
    public static bool s_CopySourceFiles = true;
    public static bool s_CopyFormattingAssembledDocx = true;
    public static bool s_ConvertUsingWord = true;

    // PowerShell oneliner that generates InlineData for all files in a directory
    // dir | % { '[InlineData("' + $_.Name + '")]' } | clip
    [Test]
    [Arguments("HC001-5DayTourPlanTemplate.docx")]
    [Arguments("HC002-Hebrew-01.docx")]
    [Arguments("HC003-Hebrew-02.docx")]
    [Arguments("HC004-ResumeTemplate.docx")]
    [Arguments("HC005-TaskPlanTemplate.docx")]
    [Arguments("HC006-Test-01.docx")]
    [Arguments("HC007-Test-02.docx")]
    [Arguments("HC008-Test-03.docx")]
    [Arguments("HC009-Test-04.docx")]
    [Arguments("HC010-Test-05.docx")]
    [Arguments("HC011-Test-06.docx")]
    [Arguments("HC012-Test-07.docx")]
    [Arguments("HC013-Test-08.docx")]
    [Arguments("HC014-RTL-Table-01.docx")]
    [Arguments("HC015-Vertical-Spacing-atLeast.docx")]
    [Arguments("HC016-Horizontal-Spacing-firstLine.docx")]
    [Arguments("HC017-Vertical-Alignment-Cell-01.docx")]
    [Arguments("HC018-Vertical-Alignment-Para-01.docx")]
    [Arguments("HC019-Hidden-Run.docx")]
    [Arguments("HC020-Small-Caps.docx")]
    [Arguments("HC021-Symbols.docx")]
    [Arguments("HC022-Table-Of-Contents.docx")]
    [Arguments("HC023-Hyperlink.docx")]
    [Arguments("HC024-Tabs-01.docx")]
    [Arguments("HC025-Tabs-02.docx")]
    [Arguments("HC026-Tabs-03.docx")]
    [Arguments("HC027-Tabs-04.docx")]
    [Arguments("HC028-No-Break-Hyphen.docx")]
    [Arguments("HC029-Table-Merged-Cells.docx")]
    [Arguments("HC030-Content-Controls.docx")]
    [Arguments("HC031-Complicated-Document.docx")]
    [Arguments("HC032-Named-Color.docx")]
    [Arguments("HC033-Run-With-Border.docx")]
    [Arguments("HC034-Run-With-Position.docx")]
    [Arguments("HC035-Strike-Through.docx")]
    [Arguments("HC036-Super-Script.docx")]
    [Arguments("HC037-Sub-Script.docx")]
    [Arguments("HC038-Conflicting-Border-Weight.docx")]
    [Arguments("HC039-Bold.docx")]
    [Arguments("HC040-Hyperlink-Fieldcode-01.docx")]
    [Arguments("HC041-Hyperlink-Fieldcode-02.docx")]
    [Arguments("HC042-Image-Png.docx")]
    [Arguments("HC043-Chart.docx")]
    [Arguments("HC044-Embedded-Workbook.docx")]
    [Arguments("HC045-Italic.docx")]
    [Arguments("HC046-BoldAndItalic.docx")]
    [Arguments("HC047-No-Section.docx")]
    [Arguments("HC048-Excerpt.docx")]
    [Arguments("HC049-Borders.docx")]
    [Arguments("HC050-Shaded-Text-01.docx")]
    [Arguments("HC051-Shaded-Text-02.docx")]
    [Arguments("HC060-Image-with-Hyperlink.docx")]
    [Arguments("HC061-Hyperlink-in-Field.docx")]
    public void HC001(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
#if COPY_FILES_FOR_DEBUGGING
        var sourceCopiedToDestDocx = new FileInfo(
            Path.Combine(TempDir, sourceDocx.Name.Replace(".docx", "-1-Source.docx"))
        );
        File.Copy(sourceDocx.FullName, sourceCopiedToDestDocx.FullName, overwrite: true);
        var assembledFormattingDestDocx = new FileInfo(
            Path.Combine(TempDir, sourceDocx.Name.Replace(".docx", "-2-FormattingAssembled.docx"))
        );
        CopyFormattingAssembledDocx(sourceDocx, assembledFormattingDestDocx);
#endif
        var oxPtConvertedDestHtml = new FileInfo(
            Path.Combine(TempDir, sourceDocx.Name.Replace(".docx", "-3-OxPt.html"))
        );
        ConvertToHtml(sourceDocx, oxPtConvertedDestHtml);
#if DO_CONVERSION_VIA_WORD
        var wordConvertedDocHtml = new FileInfo(
            Path.Combine(TestUtil.TempDir.FullName, sourceDocx.Name.Replace(".docx", "-4-Word.html"))
        );
        ConvertToHtmlUsingWord(sourceDocx, wordConvertedDocHtml);
#endif
    }

    [Test]
    [Arguments("HC006-Test-01.docx")]
    public void HC002_NoCssClasses(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var oxPtConvertedDestHtml = new FileInfo(
            Path.Combine(TempDir, sourceDocx.Name.Replace(".docx", "-5-OxPt-No-CSS-Classes.html"))
        );
        ConvertToHtmlNoCssClasses(sourceDocx, oxPtConvertedDestHtml);
    }

    private static void CopyFormattingAssembledDocx(FileInfo source, FileInfo dest)
    {
        var ba = File.ReadAllBytes(source.FullName);
        using var ms = new MemoryStream();
        ms.Write(ba, 0, ba.Length);
        using (var wordDoc = WordprocessingDocument.Open(ms, true))
        {
            RevisionAccepter.AcceptRevisions(wordDoc);
            var simplifyMarkupSettings = new SimplifyMarkupSettings
            {
                RemoveComments = true,
                RemoveContentControls = true,
                RemoveEndAndFootNotes = true,
                RemoveFieldCodes = false,
                RemoveLastRenderedPageBreak = true,
                RemovePermissions = true,
                RemoveProof = true,
                RemoveRsidInfo = true,
                RemoveSmartTags = true,
                RemoveSoftHyphens = true,
                RemoveGoBackBookmark = true,
                ReplaceTabsWithSpaces = false,
            };
            MarkupSimplifier.SimplifyMarkup(wordDoc, simplifyMarkupSettings);
            var formattingAssemblerSettings = new FormattingAssemblerSettings
            {
                RemoveStyleNamesFromParagraphAndRunProperties = false,
                ClearStyles = false,
                RestrictToSupportedLanguages = false,
                RestrictToSupportedNumberingFormats = false,
                CreateHtmlConverterAnnotationAttributes = true,
                OrderElementsPerStandard = false,
                ListItemRetrieverSettings = new ListItemRetrieverSettings()
                {
                    ListItemTextImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations,
                },
            };
            FormattingAssembler.AssembleFormatting(wordDoc, formattingAssemblerSettings);
        }

        var newBa = ms.ToArray();
        File.WriteAllBytes(dest.FullName, newBa);
    }

    private static void ConvertToHtml(FileInfo sourceDocx, FileInfo destFileName)
    {
        var byteArray = File.ReadAllBytes(sourceDocx.FullName);
        using var memoryStream = new MemoryStream();
        memoryStream.Write(byteArray, 0, byteArray.Length);
        using var wDoc = WordprocessingDocument.Open(memoryStream, true);
        var outputDirectory = destFileName.Directory;
        destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));
        var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
        var imageCounter = 0;
        var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
        if (pageTitle == null)
            pageTitle = sourceDocx.FullName;
        var settings = new WmlToHtmlConverterSettings()
        {
            PageTitle = pageTitle,
            FabricateCssClasses = true,
            CssClassPrefix = "pt-",
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false,
            ImageHandler = imageInfo =>
            {
                ++imageCounter;
                return ImageHelper.DefaultImageHandler(imageInfo, imageDirectoryName, imageCounter);
            },
        };
        var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
        // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
        // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
        // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
        // for detailed explanation.
        //
        // If you further transform the XML tree returned by ConvertToHtmlTransform, you
        // must do it correctly, or entities will not be serialized properly.
        var htmlString = html.ToString(SaveOptions.DisableFormatting);
        File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
    }

    private static void ConvertToHtmlNoCssClasses(FileInfo sourceDocx, FileInfo destFileName)
    {
        var byteArray = File.ReadAllBytes(sourceDocx.FullName);
        using var memoryStream = new MemoryStream();
        memoryStream.Write(byteArray, 0, byteArray.Length);
        using var wDoc = WordprocessingDocument.Open(memoryStream, true);
        var outputDirectory = destFileName.Directory;
        destFileName = new FileInfo(Path.Combine(outputDirectory.FullName, destFileName.Name));
        var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
        var imageCounter = 0;
        var pageTitle = (string)wDoc.CoreFilePropertiesPart.GetXDocument().Descendants(DC.title).FirstOrDefault();
        if (pageTitle == null)
            pageTitle = sourceDocx.FullName;
        var settings = new WmlToHtmlConverterSettings()
        {
            PageTitle = pageTitle,
            FabricateCssClasses = false,
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false,
            ImageHandler = imageInfo =>
            {
                ++imageCounter;
                return ImageHelper.DefaultImageHandler(imageInfo, imageDirectoryName, imageCounter);
            },
        };
        var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
        // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
        // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
        // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
        // for detailed explanation.
        //
        // If you further transform the XML tree returned by ConvertToHtmlTransform, you
        // must do it correctly, or entities will not be serialized properly.
        var htmlString = html.ToString(SaveOptions.DisableFormatting);
        File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
    }

    // Regression test for https://github.com/sergey-tihon/Clippit/issues/51
    // First tab in paragraph should not cause text overflow when text precedes the tab.
    [Test]
    public async Task HC062_FirstTabInParagraphNotIgnored()
    {
        using var memoryStream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
            settingsPart.Settings = new Settings(new DefaultTabStop { Val = 720 });

            var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
            stylesPart.Styles = new Styles();

            // Use unique, non-overlapping tokens to verify each part of the paragraph is preserved.
            var para = new Paragraph(
                new Run(new Text("BlaBlaBlaBlaBla")),
                new Run(new TabChar()),
                new Run(new Text("AfterTab"))
            );
            mainPart.Document = new Document(new Body(para));
            wordDoc.Save();
        }

        memoryStream.Position = 0;
        using var wDoc = WordprocessingDocument.Open(memoryStream, true);

        // Use inline styles (FabricateCssClasses = false) so the positioning span's
        // style rule can be inspected directly via the XElement tree.
        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            CssClassPrefix = "pt-",
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false,
        };

        var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
        var htmlString = html.ToString(SaveOptions.DisableFormatting);

        // Both unique text tokens must appear in the HTML output.
        await Assert.That(htmlString).Contains("BlaBlaBlaBlaBla");
        await Assert.That(htmlString).Contains("AfterTab");

        // Find the specific span that wraps the pre-tab text and inspect its inline style.
        // The span must use min-width (not a fixed width) so that when the preceding text
        // is wider than the tab stop, it expands instead of overflowing into subsequent content.
        var preTabSpan = html.Descendants(Xhtml.span).FirstOrDefault(s => (string)s == "BlaBlaBlaBlaBla");
        await Assert.That(preTabSpan).IsNotNull();

        var spanStyle = preTabSpan!.Attribute("style")?.Value ?? string.Empty;

        // The style rule must include min-width to allow expansion beyond the tab stop.
        await Assert.That(spanStyle).Contains("min-width:");

        // The style rule must NOT include a fixed width property, which would cause overflow.
        var styleProperties = spanStyle.Split(';').Select(p => p.Trim());
        await Assert.That(styleProperties.Where(p => p.StartsWith("width:", StringComparison.Ordinal))).IsEmpty();
    }

    [Test]
    public async Task HC063_TextBoxRenderedAsInlineBlock()
    {
        // Build a minimal DOCX that contains a floating text box (wrapSquare) with known text content.
        // The text box is anchored and should produce display:inline-block + float:left.
        using var memoryStream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

            var textBoxParagraph = BuildTextBoxParagraph(
                w,
                wp,
                a,
                wps,
                wrapElement: new XElement(wp + "wrapSquare", new XAttribute("wrapText", "bothSides")),
                text: "TextBoxContent",
                docPrId: "1"
            );

            var bodyXml = new XElement(w + "body", textBoxParagraph, new XElement(w + "p"));
            mainPart.PutXDocument(new XDocument(new XElement(w + "document", bodyXml)));
            wordDoc.Save();
        }

        memoryStream.Position = 0;
        using var wDoc = WordprocessingDocument.Open(memoryStream, true);
        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            CssClassPrefix = "pt-",
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false,
        };

        var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
        var htmlString = html.ToString(SaveOptions.DisableFormatting);

        // The text box content must appear in the HTML output
        await Assert.That(htmlString).Contains("TextBoxContent");

        // The text box must be rendered as a phrasing-safe <span> with inline-block style (not silently dropped,
        // and not as a block <div> which would be invalid HTML when nested inside <p>/<span>).
        var textBoxSpan = html.Descendants(Xhtml.span)
            .FirstOrDefault(s =>
                s.Attribute("style")?.Value?.Contains("inline-block") == true && s.Value.Contains("TextBoxContent")
            );
        await Assert.That(textBoxSpan).IsNotNull();
        var spanStyle = textBoxSpan!.Attribute("style")?.Value ?? string.Empty;
        await Assert.That(spanStyle).Contains("width:");
        await Assert.That(spanStyle).Contains("min-height:");
        await Assert.That(spanStyle).Contains("float: left");

        // Inner w:p elements must be converted to display:block <span>s, not <p>s.
        // A <p> inside a <span> is invalid HTML and causes browsers to implicitly close the outer element.
        await Assert.That(textBoxSpan.Descendants(Xhtml.p).ToList()).IsEmpty();
    }

    [Test]
    public async Task HC064_NonWrappingTextBoxHasNoFloat()
    {
        // A text box with wp:wrapNone (no text wrap / overlap) must not get float:left
        // because floating it would change layout semantics.
        using var memoryStream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

            var textBoxParagraph = BuildTextBoxParagraph(
                w,
                wp,
                a,
                wps,
                wrapElement: new XElement(wp + "wrapNone"),
                text: "TextBoxNoFloat",
                docPrId: "2"
            );

            var bodyXml = new XElement(w + "body", textBoxParagraph, new XElement(w + "p"));
            mainPart.PutXDocument(new XDocument(new XElement(w + "document", bodyXml)));
            wordDoc.Save();
        }

        memoryStream.Position = 0;
        using var wDoc = WordprocessingDocument.Open(memoryStream, true);
        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            CssClassPrefix = "pt-",
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false,
        };

        var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

        var textBoxSpan = html.Descendants(Xhtml.span)
            .FirstOrDefault(s =>
                s.Attribute("style")?.Value?.Contains("inline-block") == true && s.Value.Contains("TextBoxNoFloat")
            );
        await Assert.That(textBoxSpan).IsNotNull();
        var spanStyle = textBoxSpan!.Attribute("style")?.Value ?? string.Empty;
        await Assert.That(spanStyle).Contains("width:");
        await Assert.That(spanStyle).Contains("min-height:");
        // wrapNone means no text flow around the shape — float must not be applied
        await Assert.That(spanStyle.Contains("float")).IsFalse();
    }

    [Test]
    public async Task HC065_TextBoxWithTableContentIsNormalized()
    {
        // Regression test: w:txbxContent may contain a w:tbl, which normally converts to an HTML
        // <table>. Since the text box container is emitted as a <span>, a <table> child would produce
        // invalid HTML. The normalization must re-tag it as a display:block <span>.
        using var memoryStream = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(memoryStream, WordprocessingDocumentType.Document, true))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.AddNewPart<StyleDefinitionsPart>().Styles = new Styles();
            mainPart.AddNewPart<DocumentSettingsPart>().Settings = new Settings();

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

            // Build a text box whose content is a single-cell table instead of a paragraph
            var table = new XElement(
                w + "tbl",
                new XElement(w + "tblPr"),
                new XElement(w + "tblGrid", new XElement(w + "gridCol", new XAttribute(w + "w", "2400"))),
                new XElement(
                    w + "tr",
                    new XElement(
                        w + "tc",
                        new XElement(
                            w + "tcPr",
                            new XElement(w + "tcW", new XAttribute(w + "w", "2400"), new XAttribute(w + "type", "dxa"))
                        ),
                        new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", "CellText")))
                    )
                )
            );

            var textBoxParagraph = BuildTextBoxParagraph(
                w,
                wp,
                a,
                wps,
                wrapElement: new XElement(wp + "wrapSquare", new XAttribute("wrapText", "bothSides")),
                text: "placeholder",
                docPrId: "3"
            );

            // Replace the placeholder paragraph inside txbxContent with the table
            textBoxParagraph.Descendants(w + "txbxContent").First().ReplaceNodes(table);

            var bodyXml = new XElement(w + "body", textBoxParagraph, new XElement(w + "p"));
            mainPart.PutXDocument(new XDocument(new XElement(w + "document", bodyXml)));
            wordDoc.Save();
        }

        memoryStream.Position = 0;
        using var wDoc = WordprocessingDocument.Open(memoryStream, true);
        var settings = new WmlToHtmlConverterSettings
        {
            FabricateCssClasses = false,
            CssClassPrefix = "pt-",
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false,
        };

        var html = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

        await Assert.That(html.ToString(SaveOptions.DisableFormatting)).Contains("CellText");

        var textBoxSpan = html.Descendants(Xhtml.span)
            .FirstOrDefault(s =>
                s.Attribute("style")?.Value?.Contains("inline-block") == true && s.Value.Contains("CellText")
            );
        await Assert.That(textBoxSpan).IsNotNull();

        // No block-level HTML elements must appear as descendants of the inline text box span
        await Assert.That(textBoxSpan!.Descendants(Xhtml.p).ToList()).IsEmpty();
        await Assert.That(textBoxSpan.Descendants(Xhtml.div).ToList()).IsEmpty();
        await Assert.That(textBoxSpan.Descendants(Xhtml.table).ToList()).IsEmpty();
    }

    private static XElement BuildTextBoxParagraph(
        XNamespace w,
        XNamespace wp,
        XNamespace a,
        XNamespace wps,
        XElement wrapElement,
        string text,
        string docPrId
    )
    {
        var aNs = a.NamespaceName;
        var wpsNs = wps.NamespaceName;
        return new XElement(
            w + "p",
            new XElement(
                w + "r",
                new XElement(
                    w + "drawing",
                    new XElement(
                        wp + "anchor",
                        new XAttribute("distT", "0"),
                        new XAttribute("distB", "0"),
                        new XAttribute("distL", "114300"),
                        new XAttribute("distR", "114300"),
                        new XAttribute("simplePos", "0"),
                        new XAttribute("relativeHeight", "251658240"),
                        new XAttribute("behindDoc", "0"),
                        new XAttribute("locked", "0"),
                        new XAttribute("layoutInCell", "1"),
                        new XAttribute("allowOverlap", "1"),
                        new XElement(wp + "simplePos", new XAttribute("x", "0"), new XAttribute("y", "0")),
                        new XElement(
                            wp + "positionH",
                            new XAttribute("relativeFrom", "column"),
                            new XElement(wp + "posOffset", "0")
                        ),
                        new XElement(
                            wp + "positionV",
                            new XAttribute("relativeFrom", "paragraph"),
                            new XElement(wp + "posOffset", "0")
                        ),
                        new XElement(wp + "extent", new XAttribute("cx", "1828800"), new XAttribute("cy", "914400")),
                        wrapElement,
                        new XElement(wp + "docPr", new XAttribute("id", docPrId), new XAttribute("name", "Text Box")),
                        new XElement(
                            a + "graphic",
                            new XAttribute(XNamespace.Xmlns + "a", aNs),
                            new XElement(
                                a + "graphicData",
                                new XAttribute("uri", wpsNs),
                                new XElement(
                                    wps + "wsp",
                                    new XAttribute(XNamespace.Xmlns + "wps", wpsNs),
                                    new XElement(wps + "cNvSpPr", new XAttribute("txbx", "1")),
                                    new XElement(wps + "spPr"),
                                    new XElement(
                                        wps + "txbx",
                                        new XElement(
                                            w + "txbxContent",
                                            new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", text)))
                                        )
                                    ),
                                    new XElement(wps + "bodyPr")
                                )
                            )
                        )
                    )
                )
            )
        );
    }

#if DO_CONVERSION_VIA_WORD
    public static void ConvertToHtmlUsingWord(FileInfo sourceFileName, FileInfo destFileName)
    {
        Word.Application app = new Word.Application();
        app.Visible = false;
        try
        {
            Word.Document doc = app.Documents.Open(sourceFileName.FullName);
            doc.SaveAs2(destFileName.FullName, Word.WdSaveFormat.wdFormatFilteredHTML);
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            Console.WriteLine("Caught unexpected COM exception.");
            ((Microsoft.Office.Interop.Word._Application)app).Quit();
            Environment.Exit(0);
        }
        ((Microsoft.Office.Interop.Word._Application)app).Quit();
    }
#endif
}
