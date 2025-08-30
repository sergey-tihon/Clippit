using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word.Samples
{
    public class ReferenceAdderSamples() : Clippit.Tests.TestsBase
    {
        [Test]
        [Arguments("RaTest01.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u")] // Inserts a basic TOC before the first paragraph of the document
        [Arguments("RaTest02.docx", "/w:document/w:body/w:p[2]", @"TOC \o '1-3' \h \z \u")] // Inserts a TOC after the title of the document
        [Arguments("RaTest03.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u")] // Inserts a TOC with a different title
        [Arguments("RaTest04.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-4' \h \z \u")] // Inserts a TOC that includes headings through level 4
        [Arguments("RaTest05.docx", "/w:document/w:body/w:p[2]", @"TOC \h \z \c ""Figure""")] // Inserts a table of figures
        [Arguments("RaTest06.docx", "/w:document/w:body/w:p[1]", @"TOC \o '1-3' \h \z \u")] // Inserts a basic TOC before the first paragraph of the document. Test06.docx does not include a StylesWithEffects part.
        [Arguments("RaTest07.docx", "/w:document/w:body/w:p[2]", @"TOA \h \c ""1"" \p")] // Inserts a table of figures
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
