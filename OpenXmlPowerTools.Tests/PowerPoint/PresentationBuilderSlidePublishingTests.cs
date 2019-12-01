using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Clippit.PowerPoint;
using Xunit;

namespace Clippit.Tests.PowerPoint
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

        public PresentationBuilderSlidePublishingTests()
        {
            if (!Directory.Exists(TargetDirectory))
                Directory.CreateDirectory(TargetDirectory);
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
                slide.SaveAs(Path.Combine(targetDir, Path.GetFileName(slide.FileName)));
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

        [Theory]
        [InlineData("BRK3066.pptx")]
        public void ExtractMasters(string fileName)
        {
            var source = new PmlDocument(Path.Combine(SourceDirectory, fileName));
            int numberOfMasters;
            using (var stream = new OpenXmlMemoryStreamDocument(source))
            {
                using var doc1 = stream.GetPresentationDocument();
                numberOfMasters = doc1.PresentationPart.SlideMasterParts.Count();
            }


            var onlyMaster =
                PresentationBuilder.BuildPresentation(
                    new List<SlideSource> {new SlideSource(source, 0, 0, true)});

            onlyMaster.FileName = fileName.Replace(".pptx", "_masterOnly.pptx");
            onlyMaster.SaveAs(Path.Combine(TargetDirectory, onlyMaster.FileName));

            using var streamDoc = new OpenXmlMemoryStreamDocument(onlyMaster);
            using var resDoc = streamDoc.GetPresentationDocument();

            Assert.Empty(resDoc.PresentationPart.SlideParts);
            Assert.Equal(numberOfMasters, resDoc.PresentationPart.SlideMasterParts.Count());
        }

        [Theory]
        [InlineData("BRK3066.pptx")]
        public void ReassemblePresentationWithMaster(string fileName)
        {
            var file = Path.Combine(SourceDirectory, fileName);
            var document = new PmlDocument(file);

            // generate presentation with full master
            var onlyMaster = PresentationBuilder.BuildPresentation(
                new List<SlideSource> {new SlideSource(document, 0, 0, true)});

            // publish slides with one-layout masters
            var slides = PresentationBuilder.PublishSlides(document);

            // compose them together using only master from the first source
            var sources = new List<SlideSource> {new SlideSource(onlyMaster, true)};
            sources.AddRange(slides.Select(x => new SlideSource(x, false)));
            var newDocument = PresentationBuilder.BuildPresentation(sources);

            newDocument.FileName = fileName.Replace(".pptx", "_reassembledWithMaster.pptx");
            newDocument.SaveAs(Path.Combine(TargetDirectory, newDocument.FileName));

            var baseSize = slides.Sum(x => x.DocumentByteArray.Length);
            Assert.InRange(newDocument.DocumentByteArray.Length, 0.9 * baseSize, 1.1 * baseSize);
        }
    }
}
