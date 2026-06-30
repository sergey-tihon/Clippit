using Clippit.Cli.Infrastructure;
using Clippit.Word;

namespace Clippit.Cli.Commands.Word.Compare;

internal static class WordCompareService
{
    public static CompareResult Execute(
        InputSource source,
        InputSource revised,
        OutputTarget output,
        string? authorForRevisions,
        string? dateTimeForRevisions,
        bool caseInsensitive
    )
    {
        ValidateInputs(source, revised, output);

        var settings = new WmlComparerSettings { CaseInsensitive = caseInsensitive };
        if (!string.IsNullOrWhiteSpace(authorForRevisions))
            settings.AuthorForRevisions = authorForRevisions;
        if (!string.IsNullOrWhiteSpace(dateTimeForRevisions))
            settings.DateTimeForRevisions = dateTimeForRevisions;

        byte[] sourceBytes;
        using (var stream = source.OpenSeekable())
        using (var memory = new MemoryStream())
        {
            stream.CopyTo(memory);
            sourceBytes = memory.ToArray();
        }

        byte[] revisedBytes;
        using (var stream = revised.OpenSeekable())
        using (var memory = new MemoryStream())
        {
            stream.CopyTo(memory);
            revisedBytes = memory.ToArray();
        }

        var sourceDoc = new WmlDocument(source.LogicalName, sourceBytes);
        var revisedDoc = new WmlDocument(revised.LogicalName, revisedBytes);
        var compared = WmlComparer.Compare(sourceDoc, revisedDoc, settings);
        var revisions = WmlComparer.GetRevisions(compared, settings);

        string? tempPath = null;
        Stream outputStream;
        try
        {
            output.EnsureCanWrite(force: true, "output");
            output.EnsureDirectoryExists();
            outputStream = output.OpenWrite(out tempPath);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.OutputError($"Could not open output for writing: {ex.Message}");
        }

        try
        {
            try
            {
                outputStream.Write(compared.DocumentByteArray, 0, compared.DocumentByteArray.Length);
                outputStream.Flush();

                if (output.IsStdout)
                    output.Flush(outputStream);
            }
            finally
            {
                outputStream.Dispose();
            }

            if (!output.IsStdout)
                output.Commit(tempPath);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.OutputError($"Could not write output: {ex.Message}");
        }
        finally
        {
            OutputTarget.DeleteTemp(tempPath);
        }

        return new CompareResult
        {
            Source = source.DisplayName,
            Revised = revised.DisplayName,
            Output = output.DisplayPath,
            OutputSize = compared.DocumentByteArray.Length,
            Revisions = revisions.Count,
            AuthorForRevisions = settings.AuthorForRevisions,
            DateTimeForRevisions = settings.DateTimeForRevisions,
            CaseInsensitive = settings.CaseInsensitive,
        };
    }

    private static void ValidateInputs(InputSource source, InputSource revised, OutputTarget output)
    {
        if (source.IsStdin && revised.IsStdin)
            throw CliException.InvalidArguments("Only one input can be read from stdin.");

        if (output.IsStdout)
            return;

        if (!source.IsStdin && PathsEqual(output.DisplayPath, source.DisplayName))
            throw CliException.OutputError("Output path must not overwrite the source document.");

        if (!revised.IsStdin && PathsEqual(output.DisplayPath, revised.DisplayName))
            throw CliException.OutputError("Output path must not overwrite the revised document.");
    }

    private static bool PathsEqual(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), PathComparison);

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
}
