using System.Text.Json;
using Clippit.Cli.Infrastructure;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Word.Build;

internal static class WordBuildRunService
{
    public static WordBuildResult Execute(InputSource manifestInput, string? outputOverride, bool force)
    {
        var manifestDir = GetManifestDir(manifestInput);
        var manifest = LoadManifest(manifestInput);
        var target = OutputTarget.FromOption(outputOverride, () => ResolvePath(manifestDir, manifest.Output));

        ValidateManifest(manifest, manifestDir, target);

        target.EnsureCanWrite(force, "Output document");
        target.EnsureDirectoryExists();

        var (outputSize, entries) = BuildDocument(manifest, manifestDir, target);
        return new WordBuildResult
        {
            Output = target.DisplayPath,
            OutputSize = outputSize,
            EntryCount = entries.Count,
            Entries = entries,
        };
    }

    private static string GetManifestDir(InputSource manifestInput) =>
        manifestInput.IsStdin ? Directory.GetCurrentDirectory() : Path.GetDirectoryName(manifestInput.DisplayName)!;

    private static WordBuildManifest LoadManifest(InputSource manifestInput)
    {
        using var stream = manifestInput.OpenSeekable();
        using var reader = new StreamReader(stream);
        var manifestJson = reader.ReadToEnd();

        try
        {
            return JsonSerializer.Deserialize(manifestJson, CliJsonContext.Default.WordBuildManifest)
                ?? throw CliException.InvalidFormat("Manifest deserialized to null.");
        }
        catch (JsonException ex)
        {
            throw CliException.InvalidFormat($"Invalid manifest JSON: {ex.Message}");
        }
    }

    private static void ValidateManifest(WordBuildManifest manifest, string manifestDir, OutputTarget target)
    {
        if (manifest.Entries.Count == 0)
            throw CliException.InvalidArguments("Manifest 'entries' must contain at least one entry.");

        var hasFileEntry = false;
        foreach (var entry in manifest.Entries)
        {
            var error = entry.Validate();
            if (error is not null)
                throw CliException.InvalidArguments(error);

            if (entry.IsSection)
                continue;

            hasFileEntry = true;
            var absPath = ResolvePath(manifestDir, entry.File!);
            if (!File.Exists(absPath))
                throw CliException.FileNotFound($"Source file not found: {absPath}");

            if (!target.IsStdout && PathsEqual(target.DisplayPath, absPath))
                throw CliException.OutputError("Output path must not overwrite a source document.");
        }

        if (!hasFileEntry)
            throw CliException.InvalidArguments("Manifest 'entries' must contain at least one file entry.");
    }

    private static string ResolvePath(string manifestDir, string path) =>
        Path.IsPathRooted(path) ? path : Path.GetFullPath(Path.Combine(manifestDir, path));

    private static bool PathsEqual(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), PathComparison);

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;

    private static (long OutputSize, IReadOnlyList<WordBuildEntryResult> Entries) BuildDocument(
        WordBuildManifest manifest,
        string manifestDir,
        OutputTarget target
    )
    {
        var sources = new List<ISource>();
        var entryResults = new List<WordBuildEntryResult>();

        foreach (var entry in manifest.Entries)
        {
            if (entry.IsSection)
            {
                entryResults.Add(new WordBuildEntryResult { Section = entry.Section });
                continue;
            }

            var absPath = ResolvePath(manifestDir, entry.File!);
            var start = entry.Start ?? 0;
            var count = entry.Count ?? int.MaxValue;
            var keepSections = entry.KeepSections ?? false;
            var discardHeaders = entry.DiscardHeadersAndFootersInKeptSections ?? false;

            byte[] fileBytes;
            try
            {
                fileBytes = File.ReadAllBytes(absPath);
            }
            catch (IOException ex) when (ex is FileNotFoundException or DirectoryNotFoundException)
            {
                throw CliException.FileNotFound($"Source file not found or could not be read: {absPath}. {ex.Message}");
            }

            var actualElements = CountBodyElements(fileBytes, start, count);
            var wmlDoc = new WmlDocument(absPath, fileBytes);
            sources.Add(
                new Source(wmlDoc, start, count, keepSections)
                {
                    DiscardHeadersAndFootersInKeptSections = discardHeaders,
                }
            );
            entryResults.Add(new WordBuildEntryResult { File = entry.File, Elements = actualElements });
        }

        WmlDocument merged;
        try
        {
            merged = DocumentBuilder.BuildDocument(sources);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.InvalidFormat($"Could not build document: {ex.Message}");
        }

        string? tempPath = null;
        try
        {
            if (target.IsStdout)
            {
                using var stdout = Console.OpenStandardOutput();
                stdout.Write(merged.DocumentByteArray, 0, merged.DocumentByteArray.Length);
                stdout.Flush();
            }
            else
            {
                using (var outputStream = target.OpenWrite(out tempPath))
                {
                    outputStream.Write(merged.DocumentByteArray, 0, merged.DocumentByteArray.Length);
                    outputStream.Flush();
                }

                target.Commit(tempPath);
                tempPath = null;
            }
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.OutputError($"Could not write output: {ex.Message}");
        }
        finally
        {
            OutputTarget.DeleteTemp(tempPath);
        }

        return (merged.DocumentByteArray.Length, entryResults);
    }

    private static int CountBodyElements(byte[] fileBytes, int start, int count)
    {
        try
        {
            using var ms = new MemoryStream(fileBytes, writable: false);
            using var doc = WordprocessingDocument.Open(ms, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body is null)
                return 0;
            return body.ChildElements.Skip(start).Take(count).Count();
        }
        catch
        {
            return 0;
        }
    }
}
