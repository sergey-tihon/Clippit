using System.Xml.Linq;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
public sealed class PresentationBuilderMemoryTests
{
    private static readonly DirectoryInfo s_testFiles = new("../../../../TestFiles");

    [Test]
    public async Task MB001_DestinationSlideAnnotationClearedAfterFlush()
    {
        using var sourceStream = OpenSampleDeck();
        using var sourceDocument = PresentationDocument.Open(
            sourceStream,
            false,
            new OpenSettings { AutoSave = false }
        );
        var sourceSlidePart = GetFirstSlidePart(sourceDocument);

        using var outputStream = new MemoryStream();
        using var outputDocument = PresentationBuilder.NewDocument(outputStream);
        using (var builder = PresentationBuilder.Create(outputDocument))
        {
            var destinationSlidePart = builder.AddSlidePart(sourceSlidePart);
            destinationSlidePart.PutXDocument();
            destinationSlidePart.RemoveAnnotations<XDocument>();

            await Assert.That(destinationSlidePart.Annotation<XDocument>()).IsNull();
        }
    }

    [Test]
    public async Task MB002_SourceSlideAnnotationClearedAndRootUnloaded()
    {
        using var sourceStream = OpenSampleDeck();
        using var sourceDocument = PresentationDocument.Open(
            sourceStream,
            false,
            new OpenSettings { AutoSave = false }
        );
        var sourceSlidePart = GetFirstSlidePart(sourceDocument);

        using var outputStream = new MemoryStream();
        using var outputDocument = PresentationBuilder.NewDocument(outputStream);
        using (var builder = PresentationBuilder.Create(outputDocument))
        {
            builder.AddSlidePart(sourceSlidePart);
            sourceSlidePart.RemoveAnnotations<XDocument>();
            sourceSlidePart.UnloadRootElement();

            await Assert.That(sourceSlidePart.Annotation<XDocument>()).IsNull();
            await Assert.That(sourceSlidePart.IsRootElementLoaded).IsFalse();
        }
    }

    [Test]
    public async Task MB003_SplitStyleDestinationMutationPersistsAfterCacheClear()
    {
        using var sourceStream = OpenSampleDeck();
        using var sourceDocument = PresentationDocument.Open(
            sourceStream,
            false,
            new OpenSettings { AutoSave = false }
        );
        var sourceSlidePart = GetFirstSlidePart(sourceDocument);

        using var outputStream = new MemoryStream();
        using (var outputDocument = PresentationBuilder.NewDocument(outputStream))
        {
            using (var builder = PresentationBuilder.Create(outputDocument))
            {
                var destinationSlidePart = builder.AddSlidePart(sourceSlidePart);
                var slideDocument = destinationSlidePart.GetXDocument();
                slideDocument.Root?.SetAttributeValue(NoNamespace.show, "0");
                slideDocument.Root?.Attribute(NoNamespace.show)?.Remove();
                destinationSlidePart.PutXDocument();
                destinationSlidePart.RemoveAnnotations<XDocument>();

                await Assert.That(destinationSlidePart.Annotation<XDocument>()).IsNull();
            }
        }

        outputStream.Position = 0;
        using var reopenedDocument = PresentationDocument.Open(
            outputStream,
            false,
            new OpenSettings { AutoSave = false }
        );
        var reopenedSlidePart = GetFirstSlidePart(reopenedDocument);
        await Assert.That(reopenedSlidePart.GetXDocument().Root?.Attribute(NoNamespace.show)).IsNull();
    }

    private static Stream OpenSampleDeck() => File.OpenRead(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));

    private static SlidePart GetFirstSlidePart(PresentationDocument document)
    {
        ArgumentNullException.ThrowIfNull(document.PresentationPart);

        var slideRelId = PresentationBuilderTools.GetSlideIdsInOrder(document).First();
        return (SlidePart)document.PresentationPart.GetPartById(slideRelId);
    }
}
#pragma warning restore CA1707
