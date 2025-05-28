using System.Xml.Linq;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Clippit.Tests.PowerPoint;

public partial class PresentationBuilderSlidePublishingTests
{
    [Theory]
    [ClassData(typeof(PublishingTestData))]
    public async Task PublishUsingMemDocs(string sourcePath)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var targetDir = Path.Combine(TargetDirectory, fileName);
        if (Directory.Exists(targetDir))
            Directory.Delete(targetDir, true);
        Directory.CreateDirectory(targetDir);

        await using var srcStream = File.Open(sourcePath, FileMode.Open);
        var openSettings = new OpenSettings { AutoSave = false };
        using var srcDoc = OpenXmlExtensions.OpenPresentation(srcStream, false, openSettings);
        ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);

        var slideNumber = 0;
        var slidesIds = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc);
        foreach (var slideId in slidesIds)
        {
            var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slideId);
            var title = PresentationBuilderTools.GetSlideTitle(srcSlidePart.GetXElement());

            using var stream = new MemoryStream();
            using (var newDocument = PresentationBuilder.NewDocument(stream))
            {
                using (var builder = PresentationBuilder.Create(newDocument))
                {
                    var newSlidePart = builder.AddSlidePart(srcSlidePart);

                    // Remove the show attribute from the slide element (if it exists)
                    var slideDocument = newSlidePart.GetXDocument();
                    slideDocument.Root?.Attribute(NoNamespace.show)?.Remove();
                }

                // Set the title of the new presentation to the title of the slide
                newDocument.PackageProperties.Title = title;
            }

            var slideFileName = string.Concat(fileName, $"_{++slideNumber:000}.pptx");
            await using var fs = File.Create(Path.Combine(targetDir, slideFileName));
            stream.Position = 0;
            await stream.CopyToAsync(fs, TestContext.Current.CancellationToken);

            srcSlidePart.RemoveAnnotations<XDocument>();
            srcSlidePart.UnloadRootElement();
        }

        Log.WriteLine($"GC Total Memory: {GC.GetTotalMemory(false) / 1024 / 1024} MB");
    }

    [Theory]
    [ClassData(typeof(PublishingTestData))]
    public async Task MergeAllPowerPointBack(string sourcePath)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var targetDir = Path.Combine(TargetDirectory, fileName);
        if (!Directory.Exists(targetDir))
            Assert.Skip("Directory not found: " + targetDir);

        var slides = Directory.GetFiles(targetDir, "*.pptx", SearchOption.TopDirectoryOnly);
        if (slides.Length < 1)
            Assert.Skip("Not enough slides to merge.");
        Array.Sort(slides);

        // Create a memory stream from the original presentation
        using var ms = new MemoryStream();
        await using (var fs = File.OpenRead(sourcePath))
            await fs.CopyToAsync(ms, TestContext.Current.CancellationToken);

        // Use the first slide as the base document
        var setting = new OpenSettings { AutoSave = false };
        using (var baseDoc = PresentationDocument.Open(ms, true, setting))
        {
            ArgumentNullException.ThrowIfNull(baseDoc.PresentationPart);

            // Merge the remaining slides into the base document (one builder instance)
            using (var builder = PresentationBuilder.Create(baseDoc))
            {
                foreach (var path in slides)
                {
                    using var doc = PresentationDocument.Open(path, false, setting);
                    ArgumentNullException.ThrowIfNull(doc.PresentationPart);

                    // Add all slides in the correct order
                    foreach (var slidePath in PresentationBuilderTools.GetSlideIdsInOrder(doc))
                    {
                        var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slidePath);
                        builder.AddSlidePart(slidePart);
                    }
                }
            }

            baseDoc.PackageProperties.Title = $"{fileName} - Merged Deck X2";
        }

        // Save the merged document to a file
        var resultFile = Path.Combine(TargetDirectory, $"{fileName}_MergedDeckX2.pptx");
        ms.Position = 0;
        await using var resFile = File.Create(resultFile);
        await ms.CopyToAsync(resFile, TestContext.Current.CancellationToken);
    }
}
