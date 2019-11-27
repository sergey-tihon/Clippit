using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace Clippit.Tests.PresentationBuilderTests
{
    public class PresentationBuilderSlidePublishingTests
    {
        public static string SourceDirectory = "../../../../TestFiles/PublishSlides/";
        public static string TargetDirectory = "../../../../TestFiles/PublishSlides/output";

        public static IEnumerable<object[]> GetData()
        {
            var files = Directory.GetFiles(SourceDirectory, "*.pptx", SearchOption.TopDirectoryOnly);
            return files.OrderBy(x=>x).Select(path => new[] {path});
        }

        [Theory(Skip = "Produce same result as PublishUsingPublishSlides but slower")]
        [MemberData(nameof(GetData))]
        public void PublishUsingBuildPresentation(string path)
        {
            GenerateSlides("BuildPresentation", path, document =>
            {
                int slideCount;

                using (var streamSrcDoc = new OpenXmlMemoryStreamDocument(document)) {
                    using var srcDoc = streamSrcDoc.GetPresentationDocument();
                    slideCount = srcDoc.PresentationPart.Presentation.SlideIdList.ChildElements.Count;
                }

                var slides = new List<PmlDocument>(slideCount);
                for (var i = 0; i < slideCount; i++)
                {
                    var source = new SlideSource(document, i, 1, true);
                    var slide = PresentationBuilder.BuildPresentation(new List<SlideSource> { source });
                    slide.FileName = document.FileName.Replace(".pptx", $"_{i + 1:000}.pptx");
                    slides.Add(slide);
                }

                return slides;
            });
        }

        [Theory]
        [MemberData(nameof(GetData))]
        public void PublishUsingPublishSlides(string path)
        {
            GenerateSlides("PublishSlides", path, PresentationBuilder.PublishSlides);
        }

        private static void GenerateSlides(string subDirName, string sourcePath,
            Func<PmlDocument, IEnumerable<PmlDocument>> slideGenerator)
        {
            var targetDir = Path.Combine(TargetDirectory, Path.GetFileNameWithoutExtension(sourcePath), subDirName);
            if (Directory.Exists(targetDir))
                Directory.Delete(targetDir, true);
            Directory.CreateDirectory(targetDir);

            var document = new PmlDocument(sourcePath);
            foreach (var slide in slideGenerator(document))
            {
                slide.FileName = Path.Combine(targetDir, Path.GetFileName(slide.FileName));
                slide.Save();
            }
        }

        [Theory]
        [InlineData("BRK3066.pptx", 2)]
        public void ExtractOneSlide(string fileName, int slideNumber)
        {
            var file = Path.Combine(SourceDirectory, fileName);
            var document = new PmlDocument(file);

            var source = new SlideSource(document, slideNumber - 1, 1, true);
            var slide = PresentationBuilder.BuildPresentation(new List<SlideSource> { source });
            slide.FileName = document.FileName.Replace(".pptx", $"_{slideNumber:000}.pptx");

            slide.SaveAs(Path.Combine(TargetDirectory, Path.GetFileName(slide.FileName)));
        }

        [Theory]
        [InlineData("BRK3066.pptx")]
        public void ReassemblePresentation(string fileName)
        {
            var file = Path.Combine(SourceDirectory, fileName);
            var document = new PmlDocument(file);

            var slides = PresentationBuilder.PublishSlides(document);

            var sources = slides.Select(x => new SlideSource(x, true)).ToList();
            var newDocument = PresentationBuilder.BuildPresentation(sources);

            newDocument.FileName = fileName.Replace(".pptx", "_reassembled.pptx");
            newDocument.SaveAs(Path.Combine(TargetDirectory, newDocument.FileName));

            var baseSize = slides.Sum(x => x.DocumentByteArray.Length);
            Assert.InRange(newDocument.DocumentByteArray.Length, 0.9 * baseSize, 1.1* baseSize);
        }
    }
}
