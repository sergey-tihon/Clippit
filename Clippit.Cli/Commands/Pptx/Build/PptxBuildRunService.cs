using System.Text.Json;
using System.Xml.Linq;
using Clippit.Cli.Infrastructure;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Pptx.Build;

internal static class PptxBuildRunService
{
    public static BuildResult Execute(InputSource manifestInput, string? outputOverride, bool force)
    {
        var manifestDir = GetManifestDir(manifestInput);
        var manifest = LoadManifest(manifestInput);
        var target = OutputTarget.FromOption(outputOverride, () => ResolvePath(manifestDir, manifest.Output));

        ValidateManifest(manifest, manifestDir, target);

        target.EnsureCanWrite(force, "Output presentation");
        target.EnsureDirectoryExists();
        var (totalSlides, perEntry) = BuildPresentation(manifest, manifestDir, target);
        return new BuildResult
        {
            Output = target.DisplayPath,
            TotalSlides = totalSlides,
            Entries = perEntry,
        };
    }

    private static string GetManifestDir(InputSource manifestInput) =>
        manifestInput.IsStdin ? Directory.GetCurrentDirectory() : Path.GetDirectoryName(manifestInput.DisplayName)!;

    private static BuildManifest LoadManifest(InputSource manifestInput)
    {
        using var stream = manifestInput.OpenSeekable();
        using var reader = new StreamReader(stream);
        var manifestJson = reader.ReadToEnd();

        try
        {
            return JsonSerializer.Deserialize(manifestJson, CliJsonContext.Default.BuildManifest)
                ?? throw CliException.InvalidFormat("Manifest deserialized to null.");
        }
        catch (JsonException ex)
        {
            throw CliException.InvalidFormat($"Invalid manifest JSON: {ex.Message}");
        }
    }

    private static void ValidateManifest(BuildManifest manifest, string manifestDir, OutputTarget target)
    {
        var manifestSectionSeen = false;
        foreach (var entry in manifest.Deck)
        {
            var error = entry.Validate();
            if (error is not null)
                throw CliException.InvalidArguments(error);

            if (entry.IsSection)
            {
                manifestSectionSeen = true;
                continue;
            }

            if (manifestSectionSeen && entry.ShouldKeepSections)
                throw CliException.InvalidArguments(
                    "A file entry with 'keepSections' cannot follow manifest section entries. "
                        + "Move it before manifest sections or remove 'keepSections'."
                );
        }

        foreach (var entry in manifest.Deck.Where(e => e.IsFile))
        {
            var absPath = ResolvePath(manifestDir, entry.File!);
            if (!File.Exists(absPath))
                throw CliException.FileNotFound($"Source file not found: {absPath}");

            if (!target.IsStdout && PathsEqual(target.DisplayPath, absPath))
                throw CliException.OutputError("Output path must not overwrite a source presentation.");
        }
    }

    private static string ResolvePath(string manifestDir, string path) =>
        Path.IsPathRooted(path) ? path : Path.GetFullPath(Path.Combine(manifestDir, path));

    private static bool PathsEqual(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), PathComparison);

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

    private static (int TotalSlides, IReadOnlyList<BuildEntryResult> Entries) BuildPresentation(
        BuildManifest manifest,
        string manifestDir,
        OutputTarget target
    )
    {
        var totalSlides = 0;
        var entryResults = new List<BuildEntryResult>();
        var openSettings = new OpenSettings { AutoSave = false };
        var sectionManager = new PptxSectionManager();

        var outputStream = target.OpenWrite();

        using (outputStream)
        {
            using (var destination = PresentationBuilder.NewDocument(outputStream))
            {
                using (var builder = PresentationBuilder.Create(destination))
                {
                    foreach (var entry in manifest.Deck)
                    {
                        if (entry.IsSection)
                        {
                            sectionManager.AddNewSection(entry.Section);
                            entryResults.Add(new BuildEntryResult { Section = entry.Section });
                            continue;
                        }

                        var absPath = ResolvePath(manifestDir, entry.File!);
                        using var srcFs = new FileStream(absPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                        using var srcDoc = PresentationDocument.Open(srcFs, false, openSettings);
                        if (srcDoc.PresentationPart is null)
                            throw CliException.InvalidFormat($"Source file is not a presentation: {absPath}");

                        PptxSectionManager.SectionLoadSession? sectionSession = null;

                        if (entry.ShouldCopyAllMasters)
                        {
                            foreach (var masterPart in srcDoc.PresentationPart.SlideMasterParts)
                                builder.AddSlideMasterPart(masterPart);
                        }

                        if (entry.ShouldKeepSections)
                            sectionSession = sectionManager.LoadFrom(srcDoc);

                        var entrySlides = 0;
                        if (entry.ShouldCopySlides)
                        {
                            var slideIds = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc);
                            foreach (var slideRelId in slideIds)
                            {
                                var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slideRelId);
                                var dstSlidePart = builder.AddSlidePart(srcSlidePart);
                                dstSlidePart.GetXDocument().Descendants().Attributes("smtClean").Remove();
                                dstSlidePart.PutXDocument();
                                dstSlidePart.RemoveAnnotations<XDocument>();
                                var dstRelId = destination.PresentationPart!.GetIdOfPart(dstSlidePart);

                                if (sectionSession is not null)
                                    sectionSession.RemapRelId(slideRelId, dstRelId);
                                else
                                    sectionManager.AppendSlideToLastSection(dstRelId);

                                srcSlidePart.RemoveAnnotations<XDocument>();
                                srcSlidePart.UnloadRootElement();
                                totalSlides++;
                                entrySlides++;
                            }
                        }

                        entryResults.Add(new BuildEntryResult { File = entry.File, Slides = entrySlides });
                    }
                }

                sectionManager.SaveSectionsTo(destination);
                destination.PackageProperties.Title = manifest.Title;
                destination.PackageProperties.Creator = "Clippit";
                destination.PackageProperties.Modified = DateTime.UtcNow;
            }

            target.Flush(outputStream);
        }

        return (totalSlides, entryResults);
    }
}
