using Clippit.PowerPoint;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint
{
    public partial class PresentationBuilderSlidePublishingTests : Clippit.Tests.TestsBase
    {
        private const string SourceDirectory = "../../../../TestFiles/PublishSlides/";
        private const string TargetDirectory = "../../../../TestFiles/PublishSlides/output";

        public static class PublishingTestData
        {
            public static IEnumerable<Func<string>> Files()
            {
                var files = Directory.GetFiles(SourceDirectory, "*.pptx", SearchOption.TopDirectoryOnly);
                foreach (var file in files.OrderBy(x => x))
                    yield return () => file;
            }
        }

        public PresentationBuilderSlidePublishingTests()
        {
            if (!Directory.Exists(TargetDirectory))
                Directory.CreateDirectory(TargetDirectory);
        }

        [Test]
        [MethodDataSource(typeof(PublishingTestData), nameof(PublishingTestData.Files))]
        public async Task PublishUsingPublishSlides(string sourcePath)
        {
            var targetDir = Path.Combine(TargetDirectory, Path.GetFileNameWithoutExtension(sourcePath));
            if (Directory.Exists(targetDir))
                Directory.Delete(targetDir, true);
            Directory.CreateDirectory(targetDir);
            await using var srcStream = File.Open(sourcePath, FileMode.Open);
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
                await Assert.That(slideDoc.PackageProperties.Modified).IsEqualTo(modified);
                if (title.Equals(slideDoc.PackageProperties.Title))
                    sameTitle++;
            }

            await Assert.That(sameTitle).IsGreaterThanOrEqualTo(0).And.IsLessThanOrEqualTo(4);
        }

        [Test]
        [Arguments("BRK3066.pptx", 2)]
        public async Task ExtractOneSlide(string fileName, int slideNumber)
        {
            var file = Path.Combine(SourceDirectory, fileName);
            var document = new PmlDocument(file);
            var source = new SlideSource(document, slideNumber - 1, 1, true);
            var slide = PresentationBuilder.BuildPresentation([source]);
            slide.FileName = document.FileName.Replace(".pptx", $"_{slideNumber:000}.pptx");
            slide.SaveAs(Path.Combine(TargetDirectory, Path.GetFileName(slide.FileName)));
        }

        [Test]
        public async Task ExtractSlideWithExtendedChart()
        {
            var sourcePath = Path.Combine(SourceDirectory, "SlideWithExtendedChart.pptx");
            await using var srcStream = File.Open(sourcePath, FileMode.Open);
            var openSettings = new OpenSettings { AutoSave = false };
            using var srcDoc = OpenXmlExtensions.OpenPresentation(srcStream, false, openSettings);
            ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);
            await Assert
                .That(srcDoc.PresentationPart.SlideParts.SelectMany(slide => slide.ExtendedChartParts))
                .HasSingleItem();

            var slide = PresentationBuilder.PublishSlides(srcDoc, Path.GetFileName(sourcePath)).First();
            using var streamDoc = new OpenXmlMemoryStreamDocument(slide);
            using var slideDoc = streamDoc.GetPresentationDocument(openSettings);
            ArgumentNullException.ThrowIfNull(slideDoc.PresentationPart);
            await Assert
                .That(slideDoc.PresentationPart.SlideParts.Select(slide => slide.ExtendedChartParts))
                .HasSingleItem();
        }

        [Test]
        [Arguments("BRK3066.pptx")]
        public async Task ReassemblePresentation(string fileName)
        {
            var file = Path.Combine(SourceDirectory, fileName);
            var document = new PmlDocument(file);
            var slides = PresentationBuilder.PublishSlides(document);
            var sources = slides.Select(x => new SlideSource(x, true)).ToList();
            var newDocument = PresentationBuilder.BuildPresentation(sources);
            newDocument.FileName = fileName.Replace(".pptx", "_reassembled.pptx");
            newDocument.SaveAs(Path.Combine(TargetDirectory, newDocument.FileName));
            var baseSize = document.DocumentByteArray.Length;
            await Assert
                .That(newDocument.DocumentByteArray.Length)
                .IsBetween((int)(0.3 * baseSize), (int)(1.1 * baseSize));
        }

        [Test]
        [Arguments("BRK3066.pptx")]
        public async Task ExtractMasters(string fileName)
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
            await Assert.That(resDoc.PresentationPart.SlideParts).IsEmpty();
            await Assert.That(resDoc.PresentationPart.SlideMasterParts.Count()).IsBetween(1, numberOfMasters);
        }

        [Test]
        [Arguments("BRK3066.pptx")]
        public async Task ReassemblePresentationWithMaster(string fileName)
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
            await Assert
                .That(newDocument.DocumentByteArray.Length)
                .IsBetween((int)(0.5 * baseSize), (int)(1.1 * baseSize));
        }

        [Test]
        public async Task MergeAllPowerPoints()
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
