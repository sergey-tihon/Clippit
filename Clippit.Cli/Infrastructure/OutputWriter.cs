using System.Text.Json;
using System.Text.Json.Serialization.Metadata;

namespace Clippit.Cli.Infrastructure;

/// <summary>
/// Structured error payload written to stderr. Always compact JSON, regardless
/// of <c>--format</c> or <c>--quiet</c>. The process exit code (not duplicated
/// in the payload) is the canonical machine signal; <see cref="Code"/> is a
/// stable symbolic identifier suitable for programmatic dispatch.
/// </summary>
internal sealed class ErrorResult
{
    public required string Error { get; init; }
    public required string Code { get; init; }
}

internal enum OutputFormat
{
    Auto,
    Json,
    Text,
}

/// <summary>
/// Handles writing structured output to stdout and errors to stderr.
///
/// Format selection: explicit <c>--format</c> wins; otherwise JSON when
/// stdout is redirected (non-TTY), text when running interactively.
///
/// When <see cref="Quiet"/> is true the success result is suppressed entirely;
/// the process exit code remains the source of truth.
/// </summary>
internal sealed class OutputWriter(OutputFormat format, bool quiet = false)
{
    public bool Quiet { get; } = quiet;

    private bool UseJson =>
        format switch
        {
            OutputFormat.Json => true,
            OutputFormat.Text => false,
            _ => Console.IsOutputRedirected,
        };

    /// <summary>
    /// Writes a successful result. <paramref name="writeText"/> is invoked only
    /// in text mode; pass null to fall back to <c>result.ToString()</c>.
    /// </summary>
    public void WriteResult<T>(T result, JsonTypeInfo<T> typeInfo, Action<T, TextWriter>? writeText = null)
    {
        if (Quiet)
            return;

        if (UseJson)
        {
            Console.WriteLine(JsonSerializer.Serialize(result, typeInfo));
        }
        else if (writeText is not null)
        {
            writeText(result, Console.Out);
        }
        else
        {
            Console.WriteLine(result?.ToString());
        }
    }

    public static void WriteError(string message, string code)
    {
        var error = new ErrorResult { Error = message, Code = code };
        Console.Error.WriteLine(JsonSerializer.Serialize(error, CliJsonContext.Default.ErrorResult));
    }
}
