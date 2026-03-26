using System.Xml.Linq;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

public partial class PresentationBuilderSlidePublishingTests
{
    /// <summary>
    /// Regression test for https://github.com/sergey-tihon/Clippit/issues/155 —
    /// AddSlidePart must not throw KeyNotFoundException when a slide contains a p:oleObj
    /// or p:externalData element whose r:id is not found as either a part or an external relationship.
    /// </summary>
    [Test]
    [Arguments("oleObj")]
    [Arguments("externalData")]
    public async Task AddSlidePart_WithDanglingOleObjOrExternalDataRelationship_DoesNotThrow(string elementLocalName)
    {
        var sourcePath = Path.Combine(SourceDirectory, "BRK3066.pptx");
        var openSettings = new OpenSettings { AutoSave = false };

        // Copy the source file into a writable memory stream so we can inject a dangling reference.
        using var srcMemory = new MemoryStream();
        await using (var fs = File.OpenRead(sourcePath))
            await fs.CopyToAsync(srcMemory);
        srcMemory.Position = 0;

        using var srcDoc = PresentationDocument.Open(srcMemory, true, openSettings);
        ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);

        var firstSlideId = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc).First();
        var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(firstSlideId);

        // Inject the element with a relationship ID that doesn't exist as a part or external relationship.
        // This simulates a slide produced by third-party software with a dangling reference.
        XNamespace pns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var slideXDoc = srcSlidePart.GetXDocument();
        var spTree = slideXDoc.Descendants(pns + "spTree").FirstOrDefault();
        ArgumentNullException.ThrowIfNull(spTree);
        spTree.Add(new XElement(pns + elementLocalName, new XAttribute(rns + "id", "rId_dangling_999")));
        srcSlidePart.PutXDocument(slideXDoc);

        // Should not throw KeyNotFoundException
        using var destStream = new MemoryStream();
        using (var destDoc = PresentationBuilder.NewDocument(destStream))
        using (var builder = PresentationBuilder.Create(destDoc))
        {
            builder.AddSlidePart(srcSlidePart);
        }

        await Assert.That(destStream.Length).IsGreaterThan(0);
    }

    [Test]
    [MethodDataSource(typeof(PublishingTestData), nameof(PublishingTestData.Files))]
    public async Task PublishUsingMemDocs(string sourcePath, CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var targetDir = Path.Combine(TargetDirectory, fileName);
        if (Directory.Exists(targetDir))
            Directory.Delete(targetDir, true);
        Directory.CreateDirectory(targetDir);
        await using var srcStream = File.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
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
                await ValidateRelationships(newDocument);
            }

            var slideFileName = string.Concat(fileName, $"_{++slideNumber:000}.pptx");
            await using var fs = File.Create(Path.Combine(targetDir, slideFileName));
            stream.Position = 0;
            await stream.CopyToAsync(fs, cancellationToken).ConfigureAwait(false);
            srcSlidePart.RemoveAnnotations<XDocument>();
            srcSlidePart.UnloadRootElement();
        }

        Console.WriteLine($"GC Total Memory: {GC.GetTotalMemory(false) / 1024 / 1024} MB");
    }

    [Test]
    [MethodDataSource(typeof(PublishingTestData), nameof(PublishingTestData.Files))]
    public async Task MergeAllPowerPointBack(string sourcePath, CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var targetDir = Path.Combine(TargetDirectory, fileName);
        if (!Directory.Exists(targetDir))
        {
            Console.WriteLine("Directory not found: " + targetDir);
            return;
        }

        var slides = Directory.GetFiles(targetDir, "*.pptx", SearchOption.TopDirectoryOnly);
        if (slides.Length < 1)
        {
            Console.WriteLine("Not enough slides to merge.");
            return;
        }

        Array.Sort(slides);
        // Create a memory stream from the original presentation
        using var ms = new MemoryStream();
        await using (var fs = File.OpenRead(sourcePath))
        {
            await fs.CopyToAsync(ms, cancellationToken).ConfigureAwait(false);
        }

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
        await ms.CopyToAsync(resFile, cancellationToken).ConfigureAwait(false);
    }
}
