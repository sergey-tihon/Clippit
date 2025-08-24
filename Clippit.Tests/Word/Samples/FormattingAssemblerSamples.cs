using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word.Samples
{
    public class FormattingAssemblerSamples() : Clippit.Tests.TestsBase
    {
        [Test]
        [Arguments("Test01.docx")]
        [Arguments("Test02.docx")]
        public void Sample(string fileName)
        {
            var file = new FileInfo(Path.Combine("../../../Word/Samples/FormattingAssembler/", fileName));
            var newFile = new FileInfo(Path.Combine(TempDir, file.Name.Replace(".docx", "out.docx")));
            File.Copy(file.FullName, newFile.FullName, true);
            using var wDoc = WordprocessingDocument.Open(newFile.FullName, true);
            var settings = new FormattingAssemblerSettings()
            {
                ClearStyles = true,
                RemoveStyleNamesFromParagraphAndRunProperties = true,
                CreateHtmlConverterAnnotationAttributes = true,
                OrderElementsPerStandard = true,
                RestrictToSupportedLanguages = true,
                RestrictToSupportedNumberingFormats = true,
            };
            FormattingAssembler.AssembleFormatting(wDoc, settings);
        }
    }
}