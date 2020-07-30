using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Clippit.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
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

        [Theory]
        [MemberData(nameof(GetData))]
        public void PublishUsingPublishSlides(string sourcePath)
        {
            var targetDir = Path.Combine(TargetDirectory, Path.GetFileNameWithoutExtension(sourcePath));
            if (Directory.Exists(targetDir))
                Directory.Delete(targetDir, true);
            Directory.CreateDirectory(targetDir);

            using var srcStream = File.Open(sourcePath, FileMode.Open);
            var openSettings = new OpenSettings {AutoSave = false};
            using var srcDoc = OpenXmlExtensions.OpenPresentation(srcStream, true, openSettings);

            var title = srcDoc.PackageProperties.Title ?? string.Empty;
            var modified = srcDoc.PackageProperties.Modified;

            var sameTitle = 0;
            foreach (var slide in PresentationBuilder.PublishSlides(srcDoc, sourcePath))
            {
                slide.SaveAs(Path.Combine(targetDir, Path.GetFileName(slide.FileName)));

                using var streamDoc = new OpenXmlMemoryStreamDocument(slide);
                using var slideDoc = streamDoc.GetPresentationDocument(new OpenSettings {AutoSave = false});

                Assert.Equal(modified, slideDoc.PackageProperties.Modified);

                if (title.Equals(slideDoc.PackageProperties.Title))
                    sameTitle++;
            }
            Assert.InRange(sameTitle, 0,4);
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
            var presentation = new PmlDocument(file);

            // generate presentation with all masters
            var onlyMaster = PresentationBuilder.BuildPresentation(
                new List<SlideSource> {new SlideSource(presentation, 0, 0, true)});

            // publish slides with one-layout masters
            var slides = PresentationBuilder.PublishSlides(presentation);

            // compose them together using only master as the first source
            var sources = new List<SlideSource> {new SlideSource(onlyMaster, true)};
            sources.AddRange(slides.Select(x => new SlideSource(x, false)));
            var newDocument = PresentationBuilder.BuildPresentation(sources);

            newDocument.FileName = fileName.Replace(".pptx", "_reassembledWithMaster.pptx");
            newDocument.SaveAs(Path.Combine(TargetDirectory, newDocument.FileName));

            var baseSize = slides.Sum(x => x.DocumentByteArray.Length) + onlyMaster.DocumentByteArray.Length;
            Assert.InRange(newDocument.DocumentByteArray.Length, 0.85 * baseSize, 1.1 * baseSize);
        }
    }
}
