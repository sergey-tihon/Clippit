using System.Text.Json;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Infrastructure;

internal static class CommandRunner
{
    public static int Execute(Func<int> action)
    {
        try
        {
            return action();
        }
        catch (CliException ex)
        {
            return Fail(ex.Message, ex.ErrorCode, ex.ExitCode);
        }
        catch (FileNotFoundException ex)
        {
            return Fail(ex.Message, ErrorCodes.FileNotFound, ExitCodes.FileNotFound);
        }
        catch (PresentationBuilderException ex)
        {
            return Fail(ex.Message, ErrorCodes.InvalidFormat, ExitCodes.InvalidFormat);
        }
        catch (PowerToolsDocumentException ex)
        {
            return Fail(ex.Message, ErrorCodes.InvalidFormat, ExitCodes.InvalidFormat);
        }
        catch (OpenXmlPackageException ex)
        {
            return Fail(ex.Message, ErrorCodes.InvalidFormat, ExitCodes.InvalidFormat);
        }
        catch (FileFormatException ex)
        {
            return Fail(ex.Message, ErrorCodes.InvalidFormat, ExitCodes.InvalidFormat);
        }
        catch (InvalidDataException ex)
        {
            return Fail(ex.Message, ErrorCodes.InvalidFormat, ExitCodes.InvalidFormat);
        }
        catch (JsonException ex)
        {
            return Fail(ex.Message, ErrorCodes.InvalidFormat, ExitCodes.InvalidFormat);
        }
        catch (UnauthorizedAccessException ex)
        {
            return Fail(ex.Message, ErrorCodes.OutputError, ExitCodes.OutputError);
        }
        catch (IOException ex)
        {
            return Fail(ex.Message, ErrorCodes.OutputError, ExitCodes.OutputError);
        }
#pragma warning disable CA1031 // Top-level CLI boundary: every uncaught exception becomes a structured error.
        catch (Exception ex)
        {
            return Fail(ex.Message, ErrorCodes.InternalError, ExitCodes.InternalError);
        }
#pragma warning restore CA1031
    }

    private static int Fail(string message, string code, int exitCode)
    {
        OutputWriter.WriteError(message, code);
        return exitCode;
    }
}
