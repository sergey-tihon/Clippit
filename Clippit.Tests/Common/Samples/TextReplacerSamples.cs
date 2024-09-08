using System;
using System.IO;
using Clippit.Internal;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Common.Samples
{
    public class TextReplacerSamples : TestsBase
    {
        public TextReplacerSamples(ITestOutputHelper log)
            : base(log) { }

        private static string GetFilePath(string path) =>
            Path.Combine("../../../Common/Samples/TextReplacer/", path);

        [Theory]
        [InlineData("PowerPoint/Test01.pptx")]
        [InlineData("PowerPoint/Test02.pptx")]
        [InlineData("PowerPoint/Test03.pptx")]
        public void PowerPoint(string filePath)
        {
            var outFile = Path.Combine(
                TempDir,
                Path.GetFileName(filePath).Replace(".pptx", "out.pptx")
            );
            File.Copy(GetFilePath(filePath), outFile);
            using var pDoc = PresentationDocument.Open(outFile, true);
            TextReplacer.SearchAndReplace(pDoc, "Hello", "Goodbye", true);
        }

        [Fact]
        public void Word()
        {
            var di2 = new DirectoryInfo(GetFilePath("Word"));
            foreach (var file in di2.GetFiles("*.docx"))
                file.CopyTo(Path.Combine(TempDir, file.Name));

            using (
                var doc = WordprocessingDocument.Open(Path.Combine(TempDir, "Test01.docx"), true)
            )
                TextReplacer.SearchAndReplace(doc, "the", "this", false);

            try
            {
                using var doc = WordprocessingDocument.Open(
                    Path.Combine(TempDir, "Test02.docx"),
                    true
                );
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }
            catch (Exception) { }

            try
            {
                using var doc = WordprocessingDocument.Open(
                    Path.Combine(TempDir, "Test03.docx"),
                    true
                );
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            }
            catch (Exception) { }

            using (
                var doc = WordprocessingDocument.Open(Path.Combine(TempDir, "Test04.docx"), true)
            )
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (
                var doc = WordprocessingDocument.Open(Path.Combine(TempDir, "Test05.docx"), true)
            )
                TextReplacer.SearchAndReplace(doc, "is on", "is above", true);
            using (
                var doc = WordprocessingDocument.Open(Path.Combine(TempDir, "Test06.docx"), true)
            )
                TextReplacer.SearchAndReplace(doc, "the", "this", false);
            using (
                var doc = WordprocessingDocument.Open(Path.Combine(TempDir, "Test07.docx"), true)
            )
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (
                var doc = WordprocessingDocument.Open(Path.Combine(TempDir, "Test08.docx"), true)
            )
                TextReplacer.SearchAndReplace(doc, "the", "this", true);
            using (
                var doc = WordprocessingDocument.Open(Path.Combine(TempDir, "Test09.docx"), true)
            )
                TextReplacer.SearchAndReplace(
                    doc,
                    "===== Replace this text =====",
                    "***zzz***",
                    true
                );
        }
    }
}
