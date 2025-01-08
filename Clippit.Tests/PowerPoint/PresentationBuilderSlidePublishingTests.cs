using Clippit.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Clippit.Tests.PowerPoint
{
    public partial class PresentationBuilderSlidePublishingTests : TestsBase
    {
        private const string SourceDirectory = "../../../../TestFiles/PublishSlides/";
        private const string TargetDirectory = "../../../../TestFiles/PublishSlides/output";

        private class PublishingTestData : TheoryData<string>
        {
            public PublishingTestData()
            {
                var files = Directory.GetFiles(SourceDirectory, "*.pptx", SearchOption.TopDirectoryOnly);
                foreach (var file in files.OrderBy(x => x))
                    Add(file);
            }
        }

        public PresentationBuilderSlidePublishingTests(ITestOutputHelper log)
            : base(log)
        {
            if (!Directory.Exists(TargetDirectory))
                Directory.CreateDirectory(TargetDirectory);
        }

        [Theory]
        [ClassData(typeof(PublishingTestData))]
        public void PublishUsingPublishSlides(string sourcePath)
        {
            var targetDir = Path.Combine(TargetDirectory, Path.GetFileNameWithoutExtension(sourcePath));
            if (Directory.Exists(targetDir))
                Directory.Delete(targetDir, true);
            Directory.CreateDirectory(targetDir);

            using var srcStream = File.Open(sourcePath, FileMode.Open);
            var openSettings = new OpenSettings { AutoSave = false };
            using var srcDoc = OpenXmlExtensions.OpenPresentation(srcStream, false, openSettings);

            var title = srcDoc.PackageProperties.Title ?? string.Empty;
            var modified = srcDoc.PackageProperties.Modified;

            var sameTitle = 0;
            foreach (var slide in PresentationBuilder.PublishSlides(srcDoc, sourcePath))
            {
                slide.SaveAs(Path.Combine(targetDir, Path.GetFileName(slide.FileName)));

                using var streamDoc = new OpenXmlMemoryStreamDocument(slide);
                using var slideDoc = streamDoc.GetPresentationDocument(new OpenSettings { AutoSave = false });

                Assert.Equal(modified, slideDoc.PackageProperties.Modified);

                if (title.Equals(slideDoc.PackageProperties.Title))
                    sameTitle++;
            }
            Assert.InRange(sameTitle, 0, 4);
        }

        [Theory]
        [InlineData("BRK3066.pptx", 2)]
        public void ExtractOneSlide(string fileName, int slideNumber)
        {
            var file = Path.Combine(SourceDirectory, fileName);
            var document = new PmlDocument(file);

            var source = new SlideSource(document, slideNumber - 1, 1, true);
            var slide = PresentationBuilder.BuildPresentation([source]);

            slide.FileName = document.FileName.Replace(".pptx", $"_{slideNumber:000}.pptx");
            slide.SaveAs(Path.Combine(TargetDirectory, Path.GetFileName(slide.FileName)));
        }

        [Fact]
        public void ExtractSlideWithExtendedChart()
        {
            var sourcePath = Path.Combine(SourceDirectory, "SlideWithExtendedChart.pptx");
            using var srcStream = File.Open(sourcePath, FileMode.Open);
            var openSettings = new OpenSettings { AutoSave = false };
            using var srcDoc = OpenXmlExtensions.OpenPresentation(srcStream, false, openSettings);
            ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);

            var srcEmbeddingCount = srcDoc
                .PresentationPart.SlideParts.SelectMany(slide => slide.ExtendedChartParts)
                .Count();
            Assert.Equal(1, srcEmbeddingCount);

            var slide = PresentationBuilder.PublishSlides(srcDoc, Path.GetFileName(sourcePath)).First();
            using var streamDoc = new OpenXmlMemoryStreamDocument(slide);
            using var slideDoc = streamDoc.GetPresentationDocument(openSettings);
            ArgumentNullException.ThrowIfNull(slideDoc.PresentationPart);

            var slideEmbeddingCount = slideDoc
                .PresentationPart.SlideParts.Select(slide => slide.ExtendedChartParts)
                .Count();
            Assert.Equal(srcEmbeddingCount, slideEmbeddingCount);
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

            var baseSize = document.DocumentByteArray.Length;
            Assert.InRange(newDocument.DocumentByteArray.Length, 0.3 * baseSize, 1.1 * baseSize);
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

            var onlyMaster = PresentationBuilder.BuildPresentation([new(source, 0, 0, true)]);

            onlyMaster.FileName = fileName.Replace(".pptx", "_masterOnly.pptx");
            onlyMaster.SaveAs(Path.Combine(TargetDirectory, onlyMaster.FileName));

            using var streamDoc = new OpenXmlMemoryStreamDocument(onlyMaster);
            using var resDoc = streamDoc.GetPresentationDocument();
            ArgumentNullException.ThrowIfNull(resDoc.PresentationPart);

            Assert.Empty(resDoc.PresentationPart.SlideParts);
            Assert.InRange(resDoc.PresentationPart.SlideMasterParts.Count(), 1, numberOfMasters);
        }

        [Theory]
        [InlineData("BRK3066.pptx")]
        public void ReassemblePresentationWithMaster(string fileName)
        {
            var file = Path.Combine(SourceDirectory, fileName);
            var presentation = new PmlDocument(file);

            // generate presentation with all masters
            var onlyMaster = PresentationBuilder.BuildPresentation([new(presentation, 0, 0, true)]);

            // publish slides with one-layout masters
            var slides = PresentationBuilder.PublishSlides(presentation);

            // compose them together using only master as the first source
            var sources = new List<SlideSource> { new(onlyMaster, true) };
            sources.AddRange(slides.Select(x => new SlideSource(x, false)));
            var newDocument = PresentationBuilder.BuildPresentation(sources);

            newDocument.FileName = fileName.Replace(".pptx", "_reassembledWithMaster.pptx");
            newDocument.SaveAs(Path.Combine(TargetDirectory, newDocument.FileName));

            var baseSize = presentation.DocumentByteArray.Length;
            Assert.InRange(newDocument.DocumentByteArray.Length, 0.5 * baseSize, 1.1 * baseSize);
        }

        [Fact]
        public void MergeAllPowerPoints()
        {
            var root = SourceDirectory;
            var files = Directory
                .GetFiles(root, "*.pptx", SearchOption.TopDirectoryOnly)
                .Select(OpenXmlPowerToolsDocument.FromFileName)
                .Cast<PmlDocument>()
                .ToList();

            var sources = files.Select(x => new SlideSource(x, 0, 1000, true)).ToList();

            var result = PresentationBuilder.BuildPresentation(sources);
            var resultFile = Path.Combine(TempDir, "MergedDeck.pptx");
            result.SaveAs(resultFile);
        }
    }
}
