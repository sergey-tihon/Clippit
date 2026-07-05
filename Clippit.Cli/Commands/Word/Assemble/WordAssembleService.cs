using System.Xml;
using Clippit.Cli.Infrastructure;
using Clippit.Word;

namespace Clippit.Cli.Commands.Word.Assemble;

internal static class WordAssembleService
{
    public static AssembleResult Execute(InputSource template, InputSource data, OutputTarget output, bool force)
    {
        ValidateInputs(template, data, output);

        byte[] templateBytes;
        using (var stream = template.OpenSeekable())
        using (var memory = new MemoryStream())
        {
            stream.CopyTo(memory);
            templateBytes = memory.ToArray();
        }

        var xmlDoc = new XmlDocument { XmlResolver = null };
        try
        {
            using var stream = data.OpenSeekable();
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null };
            using var reader = XmlReader.Create(stream, settings);
            xmlDoc.Load(reader);
        }
        catch (XmlException ex)
        {
            throw CliException.InvalidFormat($"Data file is not valid XML: {ex.Message}");
        }

        var templateDoc = new WmlDocument(template.LogicalName, templateBytes);
        var assembled = DocumentAssembler.AssembleDocument(templateDoc, xmlDoc, out var templateError);

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
                outputStream.Write(assembled.DocumentByteArray, 0, assembled.DocumentByteArray.Length);
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

        return new AssembleResult
        {
            Template = template.DisplayName,
            Data = data.DisplayName,
            Output = output.DisplayPath,
            OutputSize = assembled.DocumentByteArray.Length,
            TemplateError = templateError,
        };
    }

    private static void ValidateInputs(InputSource template, InputSource data, OutputTarget output)
    {
        if (template.IsStdin && data.IsStdin)
            throw CliException.InvalidArguments("Only one input can be read from stdin.");

        if (output.IsStdout)
            return;

        if (!template.IsStdin && PathsEqual(output.DisplayPath, template.DisplayName))
            throw CliException.OutputError("Output path must not overwrite the template document.");

        if (!data.IsStdin && PathsEqual(output.DisplayPath, data.DisplayName))
            throw CliException.OutputError("Output path must not overwrite the XML data file.");
    }

    private static bool PathsEqual(string left, string right) =>
        string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), PathComparison);

    private static StringComparison PathComparison =>
        OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
}
