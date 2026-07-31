// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Text;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Clippit.Tests.Word;

public class WmlComparerTests2 : TestsBase
{
    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    public static bool m_OpenWord = false;

    ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

    [Test]
    [Arguments("CZ-1000", "CZ/CZ001-Plain.docx", "CZ/CZ001-Plain-Mod.docx")]
    [Arguments("CZ-1010", "CZ/CZ002-Multi-Paragraphs.docx", "CZ/CZ002-Multi-Paragraphs-Mod.docx")]
    [Arguments("CZ-1020", "CZ/CZ003-Multi-Paragraphs.docx", "CZ/CZ003-Multi-Paragraphs-Mod.docx")]
    [Arguments("CZ-1030", "CZ/CZ004-Multi-Paragraphs-in-Cell.docx", "CZ/CZ004-Multi-Paragraphs-in-Cell-Mod.docx")]
    public void CZ001_CompareTrackedInPrev(string testId, string name1, string name2)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var source1Docx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
        var source2Docx = new FileInfo(Path.Combine(sourceDir.FullName, name2));

        var thisTestTempDir = new DirectoryInfo(Path.Combine(TempDir, testId));
        if (thisTestTempDir.Exists)
            Assert.Fail("Duplicate test id???");
        else
            thisTestTempDir.Create();
        var source1CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source1Docx.Name));
        var source2CopiedToDestDocx = new FileInfo(Path.Combine(thisTestTempDir.FullName, source2Docx.Name));
        File.Copy(source1Docx.FullName, source1CopiedToDestDocx.FullName);
        File.Copy(source2Docx.FullName, source2CopiedToDestDocx.FullName);

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        if (m_OpenWord)
        {
            var source1DocxForWord = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2DocxForWord = new FileInfo(Path.Combine(sourceDir.FullName, name2));

            var source1CopiedToDestDocxForWord = new FileInfo(
                Path.Combine(thisTestTempDir.FullName, source1Docx.Name.Replace(".docx", "-For-Word.docx"))
            );
            var source2CopiedToDestDocxForWord = new FileInfo(
                Path.Combine(thisTestTempDir.FullName, source2Docx.Name.Replace(".docx", "-For-Word.docx"))
            );
            if (!source1CopiedToDestDocxForWord.Exists)
                File.Copy(source1Docx.FullName, source1CopiedToDestDocxForWord.FullName);
            if (!source2CopiedToDestDocxForWord.Exists)
                File.Copy(source2Docx.FullName, source2CopiedToDestDocxForWord.FullName);

            var wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
            WordRunner.RunWord(wordExe, source2CopiedToDestDocxForWord);
            WordRunner.RunWord(wordExe, source1CopiedToDestDocxForWord);
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        var before = source1CopiedToDestDocx.Name.Replace(".docx", "");
        var after = source2CopiedToDestDocx.Name.Replace(".docx", "");
        var docxWithRevisionsFi = new FileInfo(
            Path.Combine(thisTestTempDir.FullName, before + "-COMPARE-" + after + ".docx")
        );

        var source1Wml = new WmlDocument(source1CopiedToDestDocx.FullName);
        var source2Wml = new WmlDocument(source2CopiedToDestDocx.FullName);
        var settings = new WmlComparerSettings { DebugTempFileDi = thisTestTempDir };
        var comparedWml = WmlComparer.Compare(source1Wml, source2Wml, settings);

        ///////////////////////////
        comparedWml.SaveAs(docxWithRevisionsFi.FullName);
        using (var ms = new MemoryStream())
        {
            ms.Write(comparedWml.DocumentByteArray, 0, comparedWml.DocumentByteArray.Length);
            using (var wDoc = WordprocessingDocument.Open(ms, true))
            {
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
                if (errors.Any())
                {
                    var ind = "  ";
                    var sb = new StringBuilder();
                    foreach (var err in errors)
                    {
                        sb.Append("Error" + Environment.NewLine);
                        sb.Append(ind + "ErrorType: " + err.ErrorType + Environment.NewLine);
                        sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                        sb.Append(ind + "Part: " + err.Part.Uri + Environment.NewLine);
                        sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
                    }
                    var sbs = sb.ToString();
                    if (sbs != "")
                        Assert.Fail(sbs);
                }
            }
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        if (m_OpenWord)
        {
            var wordExe = new FileInfo(@"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE");
            WordRunner.RunWord(wordExe, docxWithRevisionsFi);
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    }

    private static async Task ValidateDocument(WmlDocument wmlToValidate)
    {
        using var ms = new MemoryStream();
        ms.Write(wmlToValidate.DocumentByteArray, 0, wmlToValidate.DocumentByteArray.Length);
        using var wDoc = WordprocessingDocument.Open(ms, true);
        var validator = new OpenXmlValidator();
        var errors = validator.Validate(wDoc).Where(e => !ExpectedErrors.Contains(e.Description));
        if (errors.Count() != 0)
        {
            var ind = "  ";
            var sb = new StringBuilder();
            foreach (var err in errors)
            {
                sb.Append("Error" + Environment.NewLine);
                sb.Append(ind + "ErrorType: " + err.ErrorType + Environment.NewLine);
                sb.Append(ind + "Description: " + err.Description + Environment.NewLine);
                sb.Append(ind + "Part: " + err.Part.Uri + Environment.NewLine);
                sb.Append(ind + "XPath: " + err.Path.XPath + Environment.NewLine);
            }
            var sbs = sb.ToString();
            await Assert.That(sbs).IsEqualTo("");
        }
    }

    public static string[] ExpectedErrors =
    [
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:allStyles' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:customStyles' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:latentStyles' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:stylesInUse' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:headingStyles' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:numberingStyles' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:tableStyles' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnRuns' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnParagraphs' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnNumbering' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:directFormattingOnTables' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:clearFormatting' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:top3HeadingStyles' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:visibleStyles' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:alternateStyleNames' attribute is not declared.",
        "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 1.",
        "The attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' has invalid value '0'. The MinInclusive constraint failed. The value must be greater than or equal to 2.",
    ];
}
