// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit.Html;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

/*******************************************************************************************
 * HtmlToWmlConverter expects the HTML to be passed as an XElement, i.e. as XML.  While the HTML test files that
 * are included in Open-Xml-PowerTools are able to be read as XML, most HTML is not able to be read as XML.
 * The best solution is to use the HtmlAgilityPack, which can parse HTML and save as XML.  The HtmlAgilityPack
 * is licensed under the Ms-PL (same as Open-Xml-PowerTools) so it is convenient to include it in your solution,
 * and thereby you can convert HTML to XML that can be processed by the HtmlToWmlConverter.
 *
 * A convenient way to get the DLL that has been checked out with HtmlToWmlConverter is to clone the repo at
 * https://github.com/EricWhiteDev/HtmlAgilityPack
 *
 * That repo contains only the DLL that has been checked out with HtmlToWmlConverter.
 *
 * Of course, you can also get the HtmlAgilityPack source and compile it to get the DLL.  You can find it at
 * http://codeplex.com/HtmlAgilityPack
 *
 * We don't include the HtmlAgilityPack in Open-Xml-PowerTools, to simplify installation.  The XUnit tests in
 * this module do not require the HtmlAgilityPack to run.
*******************************************************************************************/

namespace Clippit.Tests.Html;

public class HtmlToWmlConverterTests : TestsBase
{
    private static readonly bool s_ProduceAnnotatedHtml = true;

    // PowerShell oneliner that generates InlineData for all files in a directory
    // dir | % { '[InlineData("' + $_.Name + '")]' } | clip
    [Test]
    [Arguments("T0010.html")]
    [Arguments("T0011.html")]
    [Arguments("T0012.html")]
    [Arguments("T0013.html")]
    [Arguments("T0014.html")]
    [Arguments("T0015.html")]
    [Arguments("T0020.html")]
    [Arguments("T0030.html")]
    [Arguments("T0040.html")]
    [Arguments("T0050.html")]
    [Arguments("T0060.html")]
    [Arguments("T0070.html")]
    [Arguments("T0080.html")]
    [Arguments("T0090.html")]
    [Arguments("T0100.html")]
    [Arguments("T0110.html")]
    [Arguments("T0111.html")]
    [Arguments("T0112.html")]
    [Arguments("T0120.html")]
    [Arguments("T0130.html")]
    [Arguments("T0140.html")]
    [Arguments("T0150.html")]
    [Arguments("T0160.html")]
    [Arguments("T0170.html")]
    [Arguments("T0180.html")]
    [Arguments("T0190.html")]
    [Arguments("T0200.html")]
    [Arguments("T0210.html")]
    [Arguments("T0220.html")]
    [Arguments("T0230.html")]
    [Arguments("T0240.html")]
    [Arguments("T0250.html")]
    [Arguments("T0251.html")]
    [Arguments("T0260.html")]
    [Arguments("T0270.html")]
    [Arguments("T0280.html")]
    [Arguments("T0290.html")]
    [Arguments("T0300.html")]
    [Arguments("T0310.html")]
    [Arguments("T0320.html")]
    [Arguments("T0330.html")]
    [Arguments("T0340.html")]
    [Arguments("T0350.html")]
    [Arguments("T0360.html")]
    [Arguments("T0370.html")]
    [Arguments("T0380.html")]
    [Arguments("T0390.html")]
    [Arguments("T0400.html")]
    [Arguments("T0410.html")]
    [Arguments("T0420.html")]
    [Arguments("T0430.html")]
    [Arguments("T0431.html")]
    [Arguments("T0432.html")]
    [Arguments("T0440.html")]
    [Arguments("T0450.html")]
    [Arguments("T0460.html")]
    [Arguments("T0470.html")]
    [Arguments("T0480.html")]
    [Arguments("T0490.html")]
    [Arguments("T0500.html")]
    [Arguments("T0510.html")]
    [Arguments("T0520.html")]
    [Arguments("T0530.html")]
    [Arguments("T0540.html")]
    [Arguments("T0550.html")]
    [Arguments("T0560.html")]
    [Arguments("T0570.html")]
    [Arguments("T0580.html")]
    [Arguments("T0590.html")]
    [Arguments("T0600.html")]
    [Arguments("T0610.html")]
    [Arguments("T0620.html")]
    [Arguments("T0622.html")]
    [Arguments("T0630.html")]
    [Arguments("T0640.html")]
    [Arguments("T0650.html")]
    [Arguments("T0651.html")]
    [Arguments("T0660.html")]
    [Arguments("T0670.html")]
    [Arguments("T0680.html")]
    [Arguments("T0690.html")]
    [Arguments("T0691.html")]
    [Arguments("T0692.html")]
    [Arguments("T0700.html")]
    [Arguments("T0710.html")]
    [Arguments("T0720.html")]
    [Arguments("T0730.html")]
    [Arguments("T0740.html")]
    [Arguments("T0742.html")]
    [Arguments("T0745.html")]
    [Arguments("T0750.html")]
    [Arguments("T0760.html")]
    [Arguments("T0770.html")]
    [Arguments("T0780.html")]
    [Arguments("T0790.html")]
    [Arguments("T0791.html")]
    [Arguments("T0792.html")]
    [Arguments("T0793.html")]
    [Arguments("T0794.html")]
    [Arguments("T0795.html")]
    [Arguments("T0802.html")]
    [Arguments("T0804.html")]
    [Arguments("T0805.html")]
    [Arguments("T0810.html")]
    [Arguments("T0812.html")]
    [Arguments("T0814.html")]
    [Arguments("T0820.html")]
    [Arguments("T0821.html")]
    [Arguments("T0830.html")]
    [Arguments("T0840.html")]
    [Arguments("T0850.html")]
    [Arguments("T0851.html")]
    [Arguments("T0860.html")]
    [Arguments("T0870.html")]
    [Arguments("T0880.html")]
    [Arguments("T0890.html")]
    [Arguments("T0900.html")]
    [Arguments("T0910.html")]
    [Arguments("T0920.html")]
    [Arguments("T0921.html")]
    [Arguments("T0922.html")]
    [Arguments("T0923.html")]
    [Arguments("T0924.html")]
    [Arguments("T0925.html")]
    [Arguments("T0926.html")]
    [Arguments("T0927.html")]
    [Arguments("T0928.html")]
    [Arguments("T0929.html")]
    [Arguments("T0930.html")]
    [Arguments("T0931.html")]
    [Arguments("T0932.html")]
    [Arguments("T0933.html")]
    [Arguments("T0934.html")]
    [Arguments("T0935.html")]
    [Arguments("T0936.html")]
    [Arguments("T0940.html")]
    [Arguments("T0945.html")]
    [Arguments("T0948.html")]
    [Arguments("T0950.html")]
    [Arguments("T0955.html")]
    [Arguments("T0960.html")]
    [Arguments("T0968.html")]
    [Arguments("T0970.html")]
    [Arguments("T0980.html")]
    [Arguments("T0990.html")]
    [Arguments("T1000.html")]
    [Arguments("T1010.html")]
    [Arguments("T1020.html")]
    [Arguments("T1030.html")]
    [Arguments("T1040.html")]
    [Arguments("T1050.html")]
    [Arguments("T1060.html")]
    [Arguments("T1070.html")]
    [Arguments("T1080.html")]
    [Arguments("T1100.html")]
    [Arguments("T1110.html")]
    [Arguments("T1111.html")]
    [Arguments("T1112.html")]
    [Arguments("T1120.html")]
    [Arguments("T1130.html")]
    [Arguments("T1131.html")]
    [Arguments("T1132.html")]
    [Arguments("T1140.html")]
    [Arguments("T1150.html")]
    [Arguments("T1160.html")]
    [Arguments("T1170.html")]
    [Arguments("T1180.html")]
    [Arguments("T1190.html")]
    [Arguments("T1200.html")]
    [Arguments("T1201.html")]
    [Arguments("T1210.html")]
    [Arguments("T1220.html")]
    [Arguments("T1230.html")]
    [Arguments("T1240.html")]
    [Arguments("T1241.html")]
    [Arguments("T1242.html")]
    [Arguments("T1250.html")]
    [Arguments("T1251.html")]
    [Arguments("T1260.html")]
    [Arguments("T1270.html")]
    [Arguments("T1280.html")]
    [Arguments("T1290.html")]
    [Arguments("T1300.html")]
    [Arguments("T1310.html")]
    [Arguments("T1320.html")]
    [Arguments("T1330.html")]
    [Arguments("T1340.html")]
    [Arguments("T1350.html")]
    [Arguments("T1360.html")]
    [Arguments("T1370.html")]
    [Arguments("T1380.html")]
    [Arguments("T1390.html")]
    [Arguments("T1400.html")]
    [Arguments("T1410.html")]
    [Arguments("T1420.html")]
    [Arguments("T1430.html")]
    [Arguments("T1440.html")]
    [Arguments("T1450.html")]
    [Arguments("T1460.html")]
    [Arguments("T1470.html")]
    [Arguments("T1480.html")]
    [Arguments("T1490.html")]
    [Arguments("T1500.html")]
    [Arguments("T1510.html")]
    [Arguments("T1520.html")]
    [Arguments("T1530.html")]
    [Arguments("T1540.html")]
    [Arguments("T1550.html")]
    [Arguments("T1560.html")]
    [Arguments("T1570.html")]
    [Arguments("T1580.html")]
    [Arguments("T1590.html")]
    [Arguments("T1591.html")]
    [Arguments("T1610.html")]
    [Arguments("T1620.html")]
    [Arguments("T1630.html")]
    [Arguments("T1640.html")]
    [Arguments("T1650.html")]
    [Arguments("T1660.html")]
    [Arguments("T1670.html")]
    [Arguments("T1680.html")]
    [Arguments("T1690.html")]
    [Arguments("T1700.html")]
    [Arguments("T1710.html")]
    [Arguments("T1800.html")]
    [Arguments("T1810.html")]
    [Arguments("T1820.html")]
    [Arguments("T1830.html")]
    [Arguments("T1840.html")]
    [Arguments("T1850.html")]
    [Arguments("T1860.html")]
    [Arguments("T1870.html")]
    public async Task HW001(string name)
    {
#if false
            string[] cssFilter = new[] {
                "text-indent",
                "margin-left",
                "margin-right",
                "padding-left",
                "padding-right",
            };
#else
        string[] cssFilter = null;
#endif
#if false
            string[] htmlFilter = new[] {
                "img",
            };
#else
        string[] htmlFilter = null;
#endif
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceHtmlFi = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var sourceImageDi = new DirectoryInfo(
            Path.Combine(sourceDir.FullName, sourceHtmlFi.Name.Replace(".html", "_files"))
        );
        var destImageDi = new DirectoryInfo(Path.Combine(TempDir, sourceImageDi.Name));
        var sourceCopiedToDestHtmlFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-1-Source.html"))
        );
        var destCssFi = new FileInfo(Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-2.css")));
        var destDocxFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-3-ConvertedByHtmlToWml.docx"))
        );
        var annotatedHtmlFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-4-Annotated.txt"))
        );
        if (!sourceCopiedToDestHtmlFi.Exists)
        {
            Directory.CreateDirectory(sourceCopiedToDestHtmlFi.DirectoryName);
            File.Copy(sourceHtmlFi.FullName, sourceCopiedToDestHtmlFi.FullName);
        }

        var html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceCopiedToDestHtmlFi);
        var htmlString = html.ToString();
        if (htmlFilter != null && htmlFilter.Any())
        {
            var found = false;
            foreach (var item in htmlFilter)
            {
                if (htmlString.Contains(item))
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                sourceCopiedToDestHtmlFi.Delete();
                return;
            }
        }

        var usedAuthorCss = HtmlToWmlConverter.CleanUpCss(
            (string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style")
        );
        File.WriteAllText(destCssFi.FullName, usedAuthorCss);
        if (cssFilter != null && cssFilter.Any())
        {
            var found = false;
            foreach (var item in cssFilter)
            {
                if (usedAuthorCss.Contains(item))
                {
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                sourceCopiedToDestHtmlFi.Delete();
                destCssFi.Delete();
                return;
            }
        }

        if (sourceImageDi.Exists)
        {
            destImageDi.Create();
            foreach (var file in sourceImageDi.GetFiles())
            {
                File.Copy(file.FullName, destImageDi.FullName + "/" + file.Name);
            }
        }

        var settings = HtmlToWmlConverter.GetDefaultSettings();
        // image references in HTML files contain the path to the subdir that contains the images, so base URI is the name of the directory
        // that contains the HTML files
        settings.BaseUriForImages = Path.Combine(TempDir);
        var doc = HtmlToWmlConverter.ConvertHtmlToWml(
            defaultCss,
            usedAuthorCss,
            userCss,
            html,
            settings,
            null,
            s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null
        );
        await Assert.That(doc).IsNotNull();
        if (doc != null)
            SaveValidateAndFormatMainDocPart(destDocxFi, doc);
#if DO_CONVERSION_VIA_WORD
        var newAltChunkBeforeFi = new FileInfo(
            Path.Combine(TestUtil.TempDir.FullName, name.Replace(".html", "-5-AltChunkBefore.docx"))
        );
        var newAltChunkAfterFi = new FileInfo(
            Path.Combine(TestUtil.TempDir.FullName, name.Replace(".html", "-6-ConvertedViaWord.docx"))
        );
        WordAutomationUtilities.DoConversionViaWord(newAltChunkBeforeFi, newAltChunkAfterFi, html);
#endif
    }

    [Test]
    [Arguments("E0010.html")]
    [Arguments("E0020.html")]
    public void HW004(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceHtmlFi = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var sourceImageDi = new DirectoryInfo(
            Path.Combine(sourceDir.FullName, sourceHtmlFi.Name.Replace(".html", "_files"))
        );
        var destImageDi = new DirectoryInfo(Path.Combine(TempDir, sourceImageDi.Name));
        var sourceCopiedToDestHtmlFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-1-Source.html"))
        );
        var destCssFi = new FileInfo(Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-2.css")));
        var destDocxFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-3-ConvertedByHtmlToWml.docx"))
        );
        var annotatedHtmlFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-4-Annotated.txt"))
        );
        File.Copy(sourceHtmlFi.FullName, sourceCopiedToDestHtmlFi.FullName);
        var html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceCopiedToDestHtmlFi);
        var usedAuthorCss = HtmlToWmlConverter.CleanUpCss(
            (string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style")
        );
        File.WriteAllText(destCssFi.FullName, usedAuthorCss);
        var settings = HtmlToWmlConverter.GetDefaultSettings();
        settings.BaseUriForImages = Path.Combine(TempDir);
        Assert.Throws<OpenXmlPowerToolsException>(() =>
            HtmlToWmlConverter.ConvertHtmlToWml(
                defaultCss,
                usedAuthorCss,
                userCss,
                html,
                settings,
                null,
                s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null
            )
        );
    }

    [Test]
    [Arguments("T1880.html")]
    public async Task TestingNestedRowspan(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceHtmlFi = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var sourceImageDi = new DirectoryInfo(
            Path.Combine(sourceDir.FullName, sourceHtmlFi.Name.Replace(".html", "_files"))
        );
        var destImageDi = new DirectoryInfo(Path.Combine(TempDir, sourceImageDi.Name));
        var sourceCopiedToDestHtmlFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-1-Source.html"))
        );
        var destCssFi = new FileInfo(Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-2.css")));
        var destDocxFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-3-ConvertedByHtmlToWml.docx"))
        );
        var annotatedHtmlFi = new FileInfo(
            Path.Combine(TempDir, sourceHtmlFi.Name.Replace(".html", "-4-Annotated.txt"))
        );
        File.Copy(sourceHtmlFi.FullName, sourceCopiedToDestHtmlFi.FullName);
        var html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceCopiedToDestHtmlFi);
        var usedAuthorCss = HtmlToWmlConverter.CleanUpCss(
            (string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style")
        );
        await File.WriteAllTextAsync(destCssFi.FullName, usedAuthorCss);
        var settings = HtmlToWmlConverter.GetDefaultSettings();
        settings.BaseUriForImages = Path.Combine(TempDir);
        var doc = HtmlToWmlConverter.ConvertHtmlToWml(
            defaultCss,
            usedAuthorCss,
            userCss,
            html,
            settings,
            null,
            s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null
        );
        await Assert.That(doc).IsNotNull();
        if (doc != null)
            SaveValidateAndFormatMainDocPart(destDocxFi, doc);
    }

    private static async Task SaveValidateAndFormatMainDocPart(FileInfo destDocxFi, WmlDocument doc)
    {
        WmlDocument formattedDoc;
        doc.SaveAs(destDocxFi.FullName);
        using (var ms = new MemoryStream())
        {
            ms.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
            using (var document = WordprocessingDocument.Open(ms, true))
            {
                var xDoc = document.MainDocumentPart.GetXDocument();
                document.MainDocumentPart.PutXDocumentWithFormatting();
                var validator = new OpenXmlValidator();
                var errors = validator.Validate(document);
                var errorsString = errors.Select(e => e.Description + Environment.NewLine).StringConcatenate();
                // Assert that there were no errors in the generated document.
                await Assert.That(errorsString).IsEqualTo("");
            }

            formattedDoc = new WmlDocument(destDocxFi.FullName, ms.ToArray());
        }

        formattedDoc.SaveAs(destDocxFi.FullName);
    }

    private static readonly string defaultCss =
        @"html, address,
blockquote,
body, dd, div,
dl, dt, fieldset, form,
frame, frameset,
h1, h2, h3, h4,
h5, h6, noframes,
ol, p, ul, center,
dir, hr, menu, pre { display: block; unicode-bidi: embed }
li { display: list-item }
head { display: none }
table { display: table }
tr { display: table-row }
thead { display: table-header-group }
tbody { display: table-row-group }
tfoot { display: table-footer-group }
col { display: table-column }
colgroup { display: table-column-group }
td, th { display: table-cell }
caption { display: table-caption }
th { font-weight: bolder; text-align: center }
caption { text-align: center }
body { margin: auto; }
h1 { font-size: 2em; margin: auto; }
h2 { font-size: 1.5em; margin: auto; }
h3 { font-size: 1.17em; margin: auto; }
h4, p,
blockquote, ul,
fieldset, form,
ol, dl, dir,
menu { margin: auto }
a { color: blue; }
h5 { font-size: .83em; margin: auto }
h6 { font-size: .75em; margin: auto }
h1, h2, h3, h4,
h5, h6, b,
strong { font-weight: bolder }
blockquote { margin-left: 40px; margin-right: 40px }
i, cite, em,
var, address { font-style: italic }
pre, tt, code,
kbd, samp { font-family: monospace }
pre { white-space: pre }
button, textarea,
input, select { display: inline-block }
big { font-size: 1.17em }
small, sub, sup { font-size: .83em }
sub { vertical-align: sub }
sup { vertical-align: super }
table { border-spacing: 2px; }
thead, tbody,
tfoot { vertical-align: middle }
td, th, tr { vertical-align: inherit }
s, strike, del { text-decoration: line-through }
hr { border: 1px inset }
ol, ul, dir,
menu, dd { margin-left: 40px }
ol { list-style-type: decimal }
ol ul, ul ol,
ul ul, ol ol { margin-top: 0; margin-bottom: 0 }
u, ins { text-decoration: underline }
br:before { content: ""\A""; white-space: pre-line }
center { text-align: center }
:link, :visited { text-decoration: underline }
:focus { outline: thin dotted invert }
/* Begin bidirectionality settings (do not change) */
BDO[DIR=""ltr""] { direction: ltr; unicode-bidi: bidi-override }
BDO[DIR=""rtl""] { direction: rtl; unicode-bidi: bidi-override }
*[DIR=""ltr""] { direction: ltr; unicode-bidi: embed }
*[DIR=""rtl""] { direction: rtl; unicode-bidi: embed }

";
    private static readonly string userCss = @"";
}
