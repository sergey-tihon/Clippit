using System.Drawing;
using Clippit.Cli.Infrastructure;
using Clippit.Word;

namespace Clippit.Cli.Commands.Word.Consolidate;

internal static class WordConsolidateService
{
    private static readonly Color[] s_defaultPalette =
    [
        Color.FromArgb(0xFF, 0x00, 0x00), // red
        Color.FromArgb(0x00, 0x00, 0xFF), // blue
        Color.FromArgb(0x00, 0x80, 0x00), // green
        Color.FromArgb(0xFF, 0xA5, 0x00), // orange
        Color.FromArgb(0x80, 0x00, 0x80), // purple
    ];

    public static ConsolidateResult Execute(
        InputSource original,
        IReadOnlyList<InputSource> revisions,
        IReadOnlyList<string?> revisors,
        IReadOnlyList<string?> colors,
        OutputTarget output,
        bool force,
        string? authorForRevisions,
        string? dateTimeForRevisions,
        bool caseInsensitive,
        bool noTableConsolidation
    )
    {
        if (revisions.Count == 0)
            throw CliException.InvalidArguments("At least one revision document must be supplied.");

        if (revisors.Count > 0 && revisors.Count != revisions.Count)
            throw CliException.InvalidArguments(
                $"Number of --revisor values ({revisors.Count}) must match the number of revision files ({revisions.Count})."
            );

        if (colors.Count > 0 && colors.Count != revisions.Count)
            throw CliException.InvalidArguments(
                $"Number of --color values ({colors.Count}) must match the number of revision files ({revisions.Count})."
            );

        if (!output.IsStdout)
        {
            if (!original.IsStdin && PathsEqual(output.DisplayPath, original.DisplayName))
                throw CliException.OutputError("Output path must not overwrite the original document.");

            foreach (var rev in revisions)
            {
                if (!rev.IsStdin && PathsEqual(output.DisplayPath, rev.DisplayName))
                    throw CliException.OutputError("Output path must not overwrite a revision document.");
            }
        }

        var settings = new WmlComparerSettings { CaseInsensitive = caseInsensitive };
        if (!string.IsNullOrWhiteSpace(authorForRevisions))
            settings.AuthorForRevisions = authorForRevisions;
        if (!string.IsNullOrWhiteSpace(dateTimeForRevisions))
            settings.DateTimeForRevisions = dateTimeForRevisions;

        var consolidateSettings = new WmlComparerConsolidateSettings { ConsolidateWithTable = !noTableConsolidation };

        byte[] originalBytes;
        using (var stream = original.OpenSeekable())
        using (var memory = new MemoryStream())
        {
            stream.CopyTo(memory);
            originalBytes = memory.ToArray();
        }

        var originalDoc = new WmlDocument(original.LogicalName, originalBytes);

        var revisionInfos = new List<WmlRevisedDocumentInfo>(revisions.Count);
        var resultRevisions = new List<RevisionInfoResult>(revisions.Count);

        for (var i = 0; i < revisions.Count; i++)
        {
            var rev = revisions[i];
            byte[] revBytes;
            using (var stream = rev.OpenSeekable())
            using (var memory = new MemoryStream())
            {
                stream.CopyTo(memory);
                revBytes = memory.ToArray();
            }

            var revisor = revisors.Count > 0 ? revisors[i] : null;
            revisor ??= rev.IsStdin ? "Revisor" : Path.GetFileNameWithoutExtension(rev.DisplayName);

            var colorStr = colors.Count > 0 ? colors[i] : null;
            var color = ParseColor(colorStr, i);
            var colorHex = $"#{color.R:X2}{color.G:X2}{color.B:X2}";

            revisionInfos.Add(
                new WmlRevisedDocumentInfo
                {
                    RevisedDocument = new WmlDocument(rev.LogicalName, revBytes),
                    Revisor = revisor,
                    Color = color,
                }
            );

            resultRevisions.Add(
                new RevisionInfoResult
                {
                    File = rev.DisplayName,
                    Revisor = revisor,
                    Color = colorHex,
                }
            );
        }

        WmlDocument consolidated;
        try
        {
            consolidated = WmlComparer.Consolidate(originalDoc, revisionInfos, settings, consolidateSettings);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.InvalidFormat($"Could not consolidate document: {ex.Message}");
        }

        string? tempPath = null;
        Stream outputStream;
        try
        {
            output.EnsureCanWrite(force, "output");
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
                outputStream.Write(consolidated.DocumentByteArray, 0, consolidated.DocumentByteArray.Length);
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

        return new ConsolidateResult
        {
            Original = original.DisplayName,
            Revisions = resultRevisions,
            Output = output.DisplayPath,
            OutputSize = consolidated.DocumentByteArray.Length,
        };
    }

    private static Color ParseColor(string? colorStr, int index)
    {
        if (string.IsNullOrWhiteSpace(colorStr))
            return s_defaultPalette[index % s_defaultPalette.Length];

        var s = colorStr.TrimStart('#');
        if (s.Length == 6 && uint.TryParse(s, System.Globalization.NumberStyles.HexNumber, null, out var rgb))
            return Color.FromArgb((int)((rgb >> 16) & 0xFF), (int)((rgb >> 8) & 0xFF), (int)(rgb & 0xFF));

        throw CliException.InvalidArguments($"Invalid color value '{colorStr}'. Expected a hex color like #FF0000.");
    }

    private static bool PathsEqual(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), PathComparison);

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
}
