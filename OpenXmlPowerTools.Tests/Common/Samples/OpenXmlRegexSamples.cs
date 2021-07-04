using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Common.Samples
{
    public class OpenXmlRegexSamples : TestsBase
    {
        public OpenXmlRegexSamples(ITestOutputHelper log) : base(log)
        {
        }
        
        private static string GetFilePath(string path) =>
            Path.Combine("../../../Common/Samples/OpenXmlRegex/", path);

        [Fact]
        public void WordSample1()
        {
            var sourceDoc = new FileInfo(GetFilePath("TestDocument.docx"));
            var newDoc = new FileInfo(Path.Combine(TempDir, "Modified.docx"));
            File.Copy(sourceDoc.FullName, newDoc.FullName);

            using var wDoc = WordprocessingDocument.Open(newDoc.FullName, true);
            var xDoc = wDoc.MainDocumentPart.GetXDocument();

            // Match content (paragraph 1)
            var content = xDoc.Descendants(W.p).Take(1);
            var regex = new Regex("Video");
            var count = OpenXmlRegex.Match(content, regex);
            Log.WriteLine("Example #1 Count: {0}", count);

            // Match content, case insensitive (paragraph 1)
            content = xDoc.Descendants(W.p).Take(1);
            regex = new Regex("video", RegexOptions.IgnoreCase);
            count = OpenXmlRegex.Match(content, regex);
            Log.WriteLine("Example #2 Count: {0}", count);

            // Match content, with callback (paragraph 1)
            content = xDoc.Descendants(W.p).Take(1);
            regex = new Regex("video", RegexOptions.IgnoreCase);
            OpenXmlRegex.Match(content, regex, (element, match) =>
                Log.WriteLine("Example #3 Found value: >{0}<", match.Value));

            // Replace content, beginning of paragraph (paragraph 2)
            content = xDoc.Descendants(W.p).Skip(1).Take(1);
            regex = new Regex("^Video provides");
            count = OpenXmlRegex.Replace(content, regex, "Audio gives", null);
            Log.WriteLine("Example #4 Replaced: {0}", count);

            // Replace content, middle of paragraph (paragraph 3)
            content = xDoc.Descendants(W.p).Skip(2).Take(1);
            regex = new Regex("powerful");
            count = OpenXmlRegex.Replace(content, regex, "good", null);
            Log.WriteLine("Example #5 Replaced: {0}", count);

            // Replace content, end of paragraph (paragraph 4)
            content = xDoc.Descendants(W.p).Skip(3).Take(1);
            regex = new Regex(" [a-z.]*$");
            count = OpenXmlRegex.Replace(content, regex, " super good point!", null);
            Log.WriteLine("Example #6 Replaced: {0}", count);

            // Delete content, beginning of paragraph (paragraph 5)
            content = xDoc.Descendants(W.p).Skip(4).Take(1);
            regex = new Regex("^Video provides");
            count = OpenXmlRegex.Replace(content, regex, "", null);
            Log.WriteLine("Example #7 Deleted: {0}", count);

            // Delete content, middle of paragraph (paragraph 6)
            content = xDoc.Descendants(W.p).Skip(5).Take(1);
            regex = new Regex("powerful ");
            count = OpenXmlRegex.Replace(content, regex, "", null);
            Log.WriteLine("Example #8 Deleted: {0}", count);

            // Delete content, end of paragraph (paragraph 7)
            content = xDoc.Descendants(W.p).Skip(6).Take(1);
            regex = new Regex("[.]$");
            count = OpenXmlRegex.Replace(content, regex, "", null);
            Log.WriteLine("Example #9 Deleted: {0}", count);

            // Replace content in inserted text, same author (paragraph 8)
            content = xDoc.Descendants(W.p).Skip(7).Take(1);
            regex = new Regex("Video");
            count = OpenXmlRegex.Replace(content, regex, "Audio", null, true, "Eric White");
            Log.WriteLine("Example #10 Deleted: {0}", count);

            // Delete content in inserted text, same author (paragraph 9)
            content = xDoc.Descendants(W.p).Skip(8).Take(1);
            regex = new Regex("powerful ");
            count = OpenXmlRegex.Replace(content, regex, "", null, true, "Eric White");
            Log.WriteLine("Example #11 Deleted: {0}", count);

            // Replace content partially in inserted text, same author (paragraph 10)
            content = xDoc.Descendants(W.p).Skip(9).Take(1);
            regex = new Regex("Video provides ");
            count = OpenXmlRegex.Replace(content, regex, "Audio gives ", null, true, "Eric White");
            Log.WriteLine("Example #12 Replaced: {0}", count);

            // Delete content partially in inserted text, same author (paragraph 11)
            content = xDoc.Descendants(W.p).Skip(10).Take(1);
            regex = new Regex(" to help you prove your point");
            count = OpenXmlRegex.Replace(content, regex, "", null, true, "Eric White");
            Log.WriteLine("Example #13 Deleted: {0}", count);

            // Replace content in inserted text, different author (paragraph 12)
            content = xDoc.Descendants(W.p).Skip(11).Take(1);
            regex = new Regex("Video");
            count = OpenXmlRegex.Replace(content, regex, "Audio", null, true, "John Doe");
            Log.WriteLine("Example #14 Deleted: {0}", count);

            // Delete content in inserted text, different author (paragraph 13)
            content = xDoc.Descendants(W.p).Skip(12).Take(1);
            regex = new Regex("powerful ");
            count = OpenXmlRegex.Replace(content, regex, "", null, true, "John Doe");
            Log.WriteLine("Example #15 Deleted: {0}", count);

            // Replace content partially in inserted text, different author (paragraph 14)
            content = xDoc.Descendants(W.p).Skip(13).Take(1);
            regex = new Regex("Video provides ");
            count = OpenXmlRegex.Replace(content, regex, "Audio gives ", null, true, "John Doe");
            Log.WriteLine("Example #16 Replaced: {0}", count);

            // Delete content partially in inserted text, different author (paragraph 15)
            content = xDoc.Descendants(W.p).Skip(14).Take(1);
            regex = new Regex(" to help you prove your point");
            count = OpenXmlRegex.Replace(content, regex, "", null, true, "John Doe");
            Log.WriteLine("Example #17 Deleted: {0}", count);

            const string LeftDoubleQuotationMarks = @"[\u0022“„«»”]";
            const string Words = @"[\w\-&/]+(?:\s[\w\-&/]+)*";
            const string RightDoubleQuotationMarks = @"[\u0022”‟»«“]";

            // Replace content using replacement pattern (paragraph 16)
            content = xDoc.Descendants(W.p).Skip(15).Take(1);
            regex = new Regex($"{LeftDoubleQuotationMarks}(?<words>{Words}){RightDoubleQuotationMarks}");
            count = OpenXmlRegex.Replace(content, regex, "‘${words}’", null);
            Log.WriteLine("Example #18 Replaced: {0}", count);

            // Replace content using replacement pattern in partially inserted text (paragraph 17)
            content = xDoc.Descendants(W.p).Skip(16).Take(1);
            regex = new Regex($"{LeftDoubleQuotationMarks}(?<words>{Words}){RightDoubleQuotationMarks}");
            count = OpenXmlRegex.Replace(content, regex, "‘${words}’", null, true, "John Doe");
            Log.WriteLine("Example #19 Replaced: {0}", count);

            // Replace content using replacement pattern (paragraph 18)
            content = xDoc.Descendants(W.p).Skip(17).Take(1);
            regex = new Regex($"({LeftDoubleQuotationMarks})(video)({RightDoubleQuotationMarks})");
            count = OpenXmlRegex.Replace(content, regex, "$1audio$3", null, true, "John Doe");
            Log.WriteLine("Example #20 Replaced: {0}", count);

            // Recognize tabs (paragraph 19)
            content = xDoc.Descendants(W.p).Skip(18).Take(1);
            regex = new Regex(@"([1-9])\.\t");
            count = OpenXmlRegex.Replace(content, regex, "($1)\t", null);
            Log.WriteLine("Example #21 Replaced: {0}", count);

            // The next two examples deal with line breaks, i.e., the <w:br/> elements.
            // Note that you should use the U+000D (Carriage Return) character (i.e., '\r')
            // to match a <w:br/> (or <w:cr/>) and replace content with a <w:br/> element.
            // Depending on your platform, the end of line character(s) provided by
            // Environment.NewLine might be "\n" (Unix), "\r\n" (Windows), or "\r" (Mac).

            // Recognize tabs and insert line breaks (paragraph 20).
            content = xDoc.Descendants(W.p).Skip(19).Take(1);
            regex = new Regex($@"([1-9])\.{UnicodeMapper.HorizontalTabulation}");
            count = OpenXmlRegex.Replace(content, regex, $"Article $1{UnicodeMapper.CarriageReturn}", null);
            Log.WriteLine("Example #22 Replaced: {0}", count);

            // Recognize and remove line breaks (paragraph 21)
            content = xDoc.Descendants(W.p).Skip(20).Take(1);
            regex = new Regex($"{UnicodeMapper.CarriageReturn}");
            count = OpenXmlRegex.Replace(content, regex, " ", null);
            Log.WriteLine("Example #23 Replaced: {0}", count);

            // Remove soft hyphens (paragraph 22)
            var paras = xDoc.Descendants(W.p).Skip(21).Take(1).ToList();
            count = OpenXmlRegex.Replace(paras, new Regex($"{UnicodeMapper.SoftHyphen}"), "", null);
            count += OpenXmlRegex.Replace(paras, new Regex("use"), "no longer use", null);
            Log.WriteLine("Example #24 Replaced: {0}", count);

            // The next example deals with symbols (i.e., w:sym elements).
            // To work with symbols, you should acquire the Unicode values for the
            // symbols you wish to match or use in replacement patterns. The reason
            // is that UnicodeMapper will (a) mimic Microsoft Word in shifting the
            // Unicode values into the Unicode private use area (by adding U+F000)
            // and (b) use replacements for Unicode values that have been used in
            // conjunction with different fonts already (by adding U+E000).
            //
            // The replacement Únicode values will depend on the order in which
            // symbols are retrieved. Therefore, you should not rely on any fixed
            // assignment.
            //
            // In the example below, pencil will be represented by U+F021, whereas
            // spider (same value with different font) will be represented by U+E001.
            // If spider had been assigned first, spider would be U+F021 and pencil
            // would be U+E001.
            var oldPhone = UnicodeMapper.SymToChar("Wingdings", 40);
            var newPhone = UnicodeMapper.SymToChar("Wingdings", 41);
            var pencil = UnicodeMapper.SymToChar("Wingdings", 0x21);
            var spider = UnicodeMapper.SymToChar("Webdings", 0x21);

            // Replace or comment on symbols (paragraph 23)
            paras = xDoc.Descendants(W.p).Skip(22).Take(1).ToList();
            count = OpenXmlRegex.Replace(paras, new Regex($"{oldPhone}"), $"{newPhone} (replaced with new phone)", null);
            count += OpenXmlRegex.Replace(paras, new Regex($"({pencil})"), "$1 (same pencil)", null);
            count += OpenXmlRegex.Replace(paras, new Regex($"({spider})"), "$1 (same spider)", null);
            Log.WriteLine("Example #25 Replaced: {0}", count);

            wDoc.MainDocumentPart.PutXDocument();
        }

        [Fact]
        public void WordSample2()
        {
            var sourceDoc = new FileInfo(GetFilePath("TestDocument.docx"));
            var newDoc = new FileInfo(Path.Combine(TempDir,"Modified.docx"));
            File.Copy(sourceDoc.FullName, newDoc.FullName);

            using var wDoc = WordprocessingDocument.Open(newDoc.FullName, true);
            int count;
            var xDoc = wDoc.MainDocumentPart.GetXDocument();

            var content = xDoc.Descendants(W.p);
            var regex = new Regex("[.]\x020+");
            count = OpenXmlRegex.Replace(content, regex, "." + Environment.NewLine, null);

            foreach (var para in content)
            {
                var newPara = (XElement)TransformEnvironmentNewLineToParagraph(para);
                para.ReplaceNodes(newPara.Nodes());
            }

            wDoc.MainDocumentPart.PutXDocument();
        }
        
        private static object TransformEnvironmentNewLineToParagraph(XNode node)
        {
            if (node is XElement element)
            {
                if (element.Name == W.p)
                {

                }

                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(TransformEnvironmentNewLineToParagraph));
            }
            return node;
        }
        
        [Fact]
        public void PowerPointSample()
        {
            var sourcePres = new FileInfo(GetFilePath("TestPresentation.pptx"));
            var newPres = new FileInfo(Path.Combine(TempDir, "Modified.pptx"));
            File.Copy(sourcePres.FullName, newPres.FullName);

            using var pDoc = PresentationDocument.Open(newPres.FullName, true);
            foreach (var slidePart in pDoc.PresentationPart.SlideParts)
            {
                var xDoc = slidePart.GetXDocument();

                // Replace content
                var content = xDoc.Descendants(A.p);
                var regex = new Regex("Hello");
                var count = OpenXmlRegex.Replace(content, regex, "H e l l o", null);
                Log.WriteLine("Example #18 Replaced: {0}", count);

                // If you absolutely want to preserve compatibility with PowerPoint 2007, then you will need to strip the xml:space="preserve" attribute throughout.
                // This is an issue for PowerPoint only, not Word, and for 2007 only.
                // The side-effect of this is that if a run has space at the beginning or end of it, the space will be stripped upon loading, and content/layout will be affected.
                xDoc.Descendants().Attributes(XNamespace.Xml + "space").Remove();

                slidePart.PutXDocument();
            }
        }
    }
}
