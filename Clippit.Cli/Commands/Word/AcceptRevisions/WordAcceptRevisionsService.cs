using Clippit.Cli.Infrastructure;
using Clippit.Word;

namespace Clippit.Cli.Commands.Word.AcceptRevisions;

internal static class WordAcceptRevisionsService
{
    public static ConvertResult Execute(InputSource input, OutputTarget output, bool force)
    {
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

        WmlDocument accepted;
        try
        {
            var wml = new WmlDocument(input.LogicalName, inputBytes);
            accepted = RevisionAccepter.AcceptRevisions(wml);
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
                outputStream.Write(accepted.DocumentByteArray, 0, accepted.DocumentByteArray.Length);
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
            OutputSize = accepted.DocumentByteArray.Length,
        };
    }

    private static bool PathsEqual(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), PathComparison);

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
}
