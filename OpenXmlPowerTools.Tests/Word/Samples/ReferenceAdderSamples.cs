using System.IO;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Word.Samples
{
    public class ReferenceAdderSamples : TestsBase
    {
        public ReferenceAdderSamples(ITestOutputHelper log) : base(log)
        {
        }
        
        [Theory]
        [InlineData("Test01.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u")] // Inserts a basic TOC before the first paragraph of the document
        [InlineData("Test02.docx", "/w:document/w:body/w:p[2]", @"TOC \o '1-3' \h \z \u")] // Inserts a TOC after the title of the document
        [InlineData("Test03.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u")] // Inserts a TOC with a different title
        [InlineData("Test04.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-4' \h \z \u")] // Inserts a TOC that includes headings through level 4
        [InlineData("Test05.docx", "/w:document/w:body/w:p[2]", @"TOC \h \z \c ""Figure""")] // Inserts a table of figures
        [InlineData("Test06.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u")] // Inserts a basic TOC before the first paragraph of the document. Test06.docx does not include a StylesWithEffects part.
        [InlineData("Test07.docx", "/w:document/w:body/w:p[2]", @"TOA \h \c ""1"" \p")] // Inserts a table of figures
        public void Sample(string fileName, string xPath, string switches)
        {
            var srcFile = new FileInfo(Path.Combine("../../../Word/Samples/ReferenceAdder/", fileName));
            var file = Path.Combine(TempDir, srcFile.Name);
            srcFile.CopyTo(file, true);
            
            using var wDoc = WordprocessingDocument.Open(file, true);
            ReferenceAdder.AddToc(wDoc, xPath, switches, null, null);
        }
    }
}
