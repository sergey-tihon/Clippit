using System.Xml;
using Clippit.Cli.Infrastructure;
using Clippit.Word;

namespace Clippit.Cli.Commands.Word.Assemble;

internal static class WordAssembleService
{
    public static AssembleResult Execute(InputSource template, InputSource data, OutputTarget output, bool force)
    {
        ValidateInputs(template, data);

        byte[] templateBytes;
        using (var stream = template.OpenSeekable())
        using (var memory = new MemoryStream())
        {
            stream.CopyTo(memory);
            templateBytes = memory.ToArray();
        }

        string xmlContent;
        using (var stream = data.OpenSeekable())
        using (var reader = new StreamReader(stream))
        {
            xmlContent = reader.ReadToEnd();
        }

        var xmlDoc = new XmlDocument();
        try
        {
            xmlDoc.LoadXml(xmlContent);
        }
        catch (XmlException ex)
        {
            throw CliException.InvalidFormat($"Data file is not valid XML: {ex.Message}");
        }

        var templateDoc = new WmlDocument(template.LogicalName, templateBytes);
        WmlDocument assembled;
        bool templateError;
        try
        {
            assembled = DocumentAssembler.AssembleDocument(templateDoc, xmlDoc, out templateError);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.InvalidFormat($"Failed to assemble document: {ex.Message}");
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

    private static void ValidateInputs(InputSource template, InputSource data)
    {
        if (template.IsStdin && data.IsStdin)
            throw CliException.InvalidArguments("Only one input can be read from stdin.");
    }
}
