using System.IO;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Word.Samples
{
    public class FormattingAssemblerSamples : TestsBase
    {
        public FormattingAssemblerSamples(ITestOutputHelper log)
            : base(log) { }

        [Theory]
        [InlineData("Test01.docx")]
        [InlineData("Test02.docx")]
        public void Sample(string fileName)
        {
            var file = new FileInfo(
                Path.Combine("../../../Word/Samples/FormattingAssembler/", fileName)
            );
            var newFile = new FileInfo(
                Path.Combine(TempDir, file.Name.Replace(".docx", "out.docx"))
            );
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
