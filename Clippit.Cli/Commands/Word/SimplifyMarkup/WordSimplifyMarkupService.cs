using Clippit.Cli.Infrastructure;
using Clippit.Word;

namespace Clippit.Cli.Commands.Word.SimplifyMarkup;

internal static class WordSimplifyMarkupService
{
    public static ConvertResult Execute(
        InputSource input,
        OutputTarget output,
        bool force,
        bool all,
        bool acceptRevisions,
        bool removeRsidInfo,
        bool removeMarkupForDocComp,
        bool removeComments,
        bool removeBookmarks,
        bool removeContentControls,
        bool removeEndAndFootnotes,
        bool removeFieldCodes,
        bool removeGoBackBookmark,
        bool removeHyperlinks,
        bool removeLastRenderedPageBreak,
        bool removePermissions,
        bool removeProof,
        bool removeSmartTags,
        bool removeSoftHyphens,
        bool removeWebHidden,
        bool replaceTabsWithSpaces,
        bool normalizeXml
    )
    {
        var settings = new SimplifyMarkupSettings
        {
            AcceptRevisions = all || acceptRevisions,
            RemoveRsidInfo = all || removeRsidInfo || removeMarkupForDocComp,
            RemoveMarkupForDocumentComparison = all || removeMarkupForDocComp,
            RemoveComments = all || removeComments,
            RemoveBookmarks = all || removeBookmarks,
            RemoveContentControls = all || removeContentControls,
            RemoveEndAndFootNotes = all || removeEndAndFootnotes,
            RemoveFieldCodes = all || removeFieldCodes,
            RemoveGoBackBookmark = all || removeGoBackBookmark,
            RemoveHyperlinks = all || removeHyperlinks,
            RemoveLastRenderedPageBreak = all || removeLastRenderedPageBreak,
            RemovePermissions = all || removePermissions,
            RemoveProof = all || removeProof,
            RemoveSmartTags = all || removeSmartTags,
            RemoveSoftHyphens = all || removeSoftHyphens,
            RemoveWebHidden = all || removeWebHidden,
            ReplaceTabsWithSpaces = all || replaceTabsWithSpaces,
            NormalizeXml = all || normalizeXml,
        };

        if (
            !settings.AcceptRevisions
            && !settings.RemoveMarkupForDocumentComparison
            && !settings.RemoveRsidInfo
            && !settings.RemoveComments
            && !settings.RemoveBookmarks
            && !settings.RemoveContentControls
            && !settings.RemoveEndAndFootNotes
            && !settings.RemoveFieldCodes
            && !settings.RemoveGoBackBookmark
            && !settings.RemoveHyperlinks
            && !settings.RemoveLastRenderedPageBreak
            && !settings.RemovePermissions
            && !settings.RemoveProof
            && !settings.RemoveSmartTags
            && !settings.RemoveSoftHyphens
            && !settings.RemoveWebHidden
            && !settings.ReplaceTabsWithSpaces
            && !settings.NormalizeXml
        )
        {
            throw CliException.InvalidArguments(
                "At least one simplification flag must be provided. Use --all to enable all cleanup options."
            );
        }

        if (!output.IsStdout)
        {
            if (!input.IsStdin && PathsEqual(output.DisplayPath, input.DisplayName))
                throw CliException.OutputError("Output path must not overwrite the input document.");
        }

        byte[] inputBytes;
        using (var stream = input.OpenSeekable())
        using (var memory = new MemoryStream())
        {
            stream.CopyTo(memory);
            inputBytes = memory.ToArray();
        }

        WmlDocument simplified;
        try
        {
            var wml = new WmlDocument(input.LogicalName, inputBytes);
            simplified = MarkupSimplifier.SimplifyMarkup(wml, settings);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.InvalidFormat($"Could not process document: {ex.Message}");
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
                outputStream.Write(simplified.DocumentByteArray, 0, simplified.DocumentByteArray.Length);
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

        return new ConvertResult
        {
            Input = input.DisplayName,
            Output = output.DisplayPath,
            OutputSize = simplified.DocumentByteArray.Length,
        };
    }

    private static bool PathsEqual(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), PathComparison);

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
}
