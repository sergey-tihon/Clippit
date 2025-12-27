// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
#define COPY_FILES_FOR_DEBUGGING
// DO_CONVERSION_VIA_WORD is defined in the project Clippit.Tests.OA.csproj, but not in the Clippit.Tests.csproj
using System.Text;
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

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
