using Clippit.PowerPoint;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

/// <summary>
/// Tests for the file-based <see cref="PresentationBuilder.PublishSlides(string,string)"/> and
/// <see cref="PresentationBuilder.PublishSlides(PresentationDocument,string,string)"/> overloads
/// that write each slide directly to disk, avoiding large heap allocations.
/// </summary>
public partial class PresentationBuilderSlidePublishingTests
{
    [Test]
    [Arguments("BRK3066.pptx")]
    public async Task PublishSlides_FilePath_CreatesSlideFiles(string fileName)
    {
        var srcPath = Path.Combine(SourceDirectory, fileName);
        var outputDir = Path.Combine(TempDir, "FileBased_" + Path.GetFileNameWithoutExtension(fileName));

        var outputPaths = PresentationBuilder.PublishSlides(srcPath, outputDir).ToList();

        await Assert.That(outputPaths).IsNotEmpty();

        foreach (var path in outputPaths)
        {
            await Assert.That(File.Exists(path)).IsTrue();
            // Verify each output is a valid presentation
            using var doc = PresentationDocument.Open(path, false);
            await Assert.That(doc.PresentationPart).IsNotNull();
            await Assert.That(doc.PresentationPart!.SlideParts).HasSingleItem();
        }
    }

    [Test]
    [Arguments("BRK3066.pptx")]
    public async Task PublishSlides_FilePath_ProducesSameCountAsPmlOverload(string fileName)
    {
        var srcPath = Path.Combine(SourceDirectory, fileName);
        var outputDir = Path.Combine(TempDir, "FileBasedCount_" + Path.GetFileNameWithoutExtension(fileName));

        var filePaths = PresentationBuilder.PublishSlides(srcPath, outputDir).ToList();
        var pmlSlides = PresentationBuilder.PublishSlides(new PmlDocument(srcPath)).ToList();

        await Assert.That(filePaths.Count).IsEqualTo(pmlSlides.Count);
    }

    [Test]
    [Arguments("BRK3066.pptx")]
    public async Task PublishSlides_FilePath_CreatesOutputDirectory(string fileName)
    {
        var srcPath = Path.Combine(SourceDirectory, fileName);
        var outputDir = Path.Combine(TempDir, "AutoCreate_" + Guid.NewGuid().ToString("N"));

        await Assert.That(Directory.Exists(outputDir)).IsFalse();

        var outputPaths = PresentationBuilder.PublishSlides(srcPath, outputDir).ToList();

        await Assert.That(Directory.Exists(outputDir)).IsTrue();
        await Assert.That(outputPaths).IsNotEmpty();
    }

    [Test]
    [Arguments("BRK3066.pptx")]
    public async Task PublishSlides_OpenDocument_WritesToFiles(string fileName)
    {
        var srcPath = Path.Combine(SourceDirectory, fileName);
        var outputDir = Path.Combine(TempDir, "OpenDoc_" + Path.GetFileNameWithoutExtension(fileName));

        using var srcDoc = PresentationDocument.Open(srcPath, isEditable: false);
        var outputPaths = PresentationBuilder.PublishSlides(srcDoc, fileName, outputDir).ToList();

        await Assert.That(outputPaths).IsNotEmpty();

        foreach (var path in outputPaths)
        {
            await Assert.That(File.Exists(path)).IsTrue();
            using var doc = PresentationDocument.Open(path, false);
            await Assert.That(doc.PresentationPart).IsNotNull();
            await Assert.That(doc.PresentationPart!.SlideParts).HasSingleItem();
        }
    }

    [Test]
    [Arguments("BRK3066.pptx")]
    public async Task PublishSlides_FilePath_OutputMatchesPmlContent(string fileName)
    {
        var srcPath = Path.Combine(SourceDirectory, fileName);
        var outputDir = Path.Combine(TempDir, "ContentCheck_" + Path.GetFileNameWithoutExtension(fileName));

        var filePaths = PresentationBuilder.PublishSlides(srcPath, outputDir).ToList();
        var pmlSlides = PresentationBuilder.PublishSlides(new PmlDocument(srcPath)).ToList();

        await Assert.That(filePaths.Count).IsEqualTo(pmlSlides.Count);

        for (var i = 0; i < filePaths.Count; i++)
        {
            using var fileDoc = PresentationDocument.Open(filePaths[i], false);
            using var pmlStream = new OpenXmlMemoryStreamDocument(pmlSlides[i]);
            using var pmlDoc = pmlStream.GetPresentationDocument(new OpenSettings { AutoSave = false });

            await Assert.That(fileDoc.PackageProperties.Title).IsEqualTo(pmlDoc.PackageProperties.Title);
        }
    }
}
