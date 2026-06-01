using System.Diagnostics.CodeAnalysis;

namespace Clippit.Cli.Infrastructure;

[SuppressMessage(
    "Design",
    "CA1032:Implement standard exception constructors",
    Justification = "CLI exceptions must carry a stable error code and process exit code."
)]
internal sealed class CliException(string message, string errorCode, int exitCode) : Exception(message)
{
    public string ErrorCode { get; } = errorCode;

    public int ExitCode { get; } = exitCode;

    public static CliException InvalidArguments(string message) =>
        new(message, ErrorCodes.InvalidArguments, ExitCodes.InvalidArguments);

    public static CliException InvalidFormat(string message) =>
        new(message, ErrorCodes.InvalidFormat, ExitCodes.InvalidFormat);

    public static CliException OutputError(string message) =>
        new(message, ErrorCodes.OutputError, ExitCodes.OutputError);

    public static CliException FileNotFound(string message) =>
        new(message, ErrorCodes.FileNotFound, ExitCodes.FileNotFound);
}
