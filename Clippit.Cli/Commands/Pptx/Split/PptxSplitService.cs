using System.Globalization;
using System.Text.Json;
using System.Xml.Linq;
using Clippit.Cli.Commands.Pptx.Build;
using Clippit.Cli.Infrastructure;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Pptx.Split;

internal static class PptxSplitService
{
    public static SplitResult Execute(
        InputSource input,
        DirectoryInfo outputDir,
        string? slidesExpression,
        bool force,
        bool writeManifest
    )
    {
        var inputStream = input.OpenSeekable();
        using (inputStream)
        using (
            var sourceDocument = PresentationDocument.Open(inputStream, false, new OpenSettings { AutoSave = false })
        )
        {
            if (sourceDocument.PresentationPart is null)
                throw CliException.InvalidFormat("Not a Presentation document.");

            var slideIds = PresentationBuilderTools.GetSlideIdsInOrder(sourceDocument);
            var indexes = SlideSelectionParser.Parse(slidesExpression, slideIds.Count);
            var width = Math.Max(3, slideIds.Count.ToString(CultureInfo.InvariantCulture).Length);

            var baseName = GetSourceBaseName(input.LogicalName);
            var manifestPath = writeManifest ? Path.Combine(outputDir.FullName, $"{baseName}.manifest.json") : null;

            if (!force)
            {
                foreach (var index in indexes)
                {
                    var path = Path.Combine(outputDir.FullName, GetSlideFileName(input.LogicalName, index, width));
                    if (File.Exists(path))
                        throw CliException.OutputError(
                            $"Output file already exists: {path}. Pass --force to overwrite."
                        );
                }

                if (manifestPath is not null && File.Exists(manifestPath))
                    throw CliException.OutputError(
                        $"Output file already exists: {manifestPath}. Pass --force to overwrite."
                    );
            }

            var slides = indexes
                .Select(index => WriteSlide(sourceDocument, slideIds, outputDir, input.LogicalName, index, width))
                .ToList();

            string? manifestDisplay = null;
            if (manifestPath is not null)
            {
                WriteManifest(sourceDocument, slideIds, indexes, slides, baseName, manifestPath);
                manifestDisplay = manifestPath;
            }

            return new SplitResult
            {
                Input = input.DisplayName,
                OutputDir = outputDir.FullName,
                Manifest = manifestDisplay,
                Slides = slides,
                Count = slides.Count,
            };
        }
    }

    private static string GetSourceBaseName(string inputFileName)
    {
        var extension = Path.GetExtension(inputFileName);
        return extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase)
            ? Path.GetFileNameWithoutExtension(inputFileName)
            : inputFileName;
    }

    private static void WriteManifest(
        PresentationDocument sourceDocument,
        List<string> slideIds,
        List<int> indexes,
        List<SlideEntry> slides,
        string baseName,
        string manifestPath
    )
    {
        var sourceSections = PptxSectionManager.ReadSections(sourceDocument);
        Dictionary<string, string>? slideToSection = null;
        if (sourceSections.Count > 0)
        {
            slideToSection = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var (name, _, relIds) in sourceSections)
            foreach (var relId in relIds)
                slideToSection[relId] = name;
        }

        var entries = new List<DeckEntry>();
        string? currentSection = null;
        var manifestDir = Path.GetDirectoryName(manifestPath)!;

        for (var i = 0; i < indexes.Count; i++)
        {
            var slideIndex = indexes[i];
            var slideRelId = slideIds[slideIndex - 1];

            if (slideToSection is not null && slideToSection.TryGetValue(slideRelId, out var section))
            {
                if (!string.Equals(section, currentSection, StringComparison.Ordinal))
                {
                    entries.Add(new DeckEntry { Section = section });
                    currentSection = section;
                }
            }

            var slideFile = slides[i].File;
            var relative = Path.GetRelativePath(manifestDir, slideFile);
            entries.Add(new DeckEntry { File = relative });
        }

        var manifest = new BuildManifest
        {
            Schema = CliConstants.DeckManifestSchema,
            Title = baseName,
            Output = $"{baseName}.merged.pptx",
            Deck = entries,
        };

        var json = JsonSerializer.Serialize(manifest, CliJsonContextIndented.Default.BuildManifest);
        File.WriteAllText(manifestPath, json);
    }

    private static SlideEntry WriteSlide(
        PresentationDocument sourceDocument,
        List<string> slideIds,
        DirectoryInfo outputDir,
        string inputFileName,
        int index,
        int indexWidth
    )
    {
        var slideRelId = slideIds[index - 1];
        var sourceSlidePart = (SlidePart)sourceDocument.PresentationPart!.GetPartById(slideRelId);
        var destPath = Path.Combine(outputDir.FullName, GetSlideFileName(inputFileName, index, indexWidth));
        var title = WriteSingleSlide(sourceSlidePart, destPath);
        sourceSlidePart.RemoveAnnotations<XDocument>();
        sourceSlidePart.UnloadRootElement();

        return new SlideEntry
        {
            Index = index,
            File = destPath,
            Title = title,
        };
    }

    private static string GetSlideFileName(string inputFileName, int slideIndex, int width)
    {
        var baseName = GetSourceBaseName(inputFileName);
        var indexText = slideIndex.ToString(CultureInfo.InvariantCulture).PadLeft(width, '0');
        return $"{baseName}_{indexText}.pptx";
    }

    private static string? WriteSingleSlide(SlidePart sourceSlidePart, string destPath)
    {
        using var outputStream = new FileStream(destPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None);
        using var outputDocument = PresentationBuilder.NewDocument(outputStream);
        using (var builder = PresentationBuilder.Create(outputDocument))
        {
            var newSlidePart = builder.AddSlidePart(sourceSlidePart);
            var slideDocument = newSlidePart.GetXDocument();
            slideDocument.Root?.Attribute(NoNamespace.show)?.Remove();
            slideDocument.Descendants().Attributes("smtClean").Remove();
            newSlidePart.PutXDocument();
            newSlidePart.RemoveAnnotations<XDocument>();
        }

        var title = PresentationBuilderTools.GetSlideTitle(sourceSlidePart.GetXElement());
        outputDocument.PackageProperties.Title = title;
        return title.Length > 0 ? title : null;
    }
}
