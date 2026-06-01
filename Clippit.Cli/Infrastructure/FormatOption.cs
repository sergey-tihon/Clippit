using System.CommandLine;

namespace Clippit.Cli.Infrastructure;

/// <summary>
/// Shared factory for the global <c>--format</c> / <c>-f</c> option and the
/// <c>--quiet</c> / <c>-q</c> option.
///
/// Output discipline (applies to every command that emits a result):
///   - <c>--format json</c>  : always JSON to stdout
///   - <c>--format text</c>  : always human text to stdout
///   - (unset)               : auto — JSON when stdout is redirected, text on TTY
///   - <c>--quiet</c>        : suppresses the success summary on stdout. Exit code
///                             still reflects success/failure. Errors are unaffected.
///
/// Errors are always emitted as a single compact JSON object on stderr,
/// regardless of <c>--format</c> or <c>--quiet</c>.
/// </summary>
internal static class FormatOption
{
    public static Option<OutputFormat> BuildFormatOption()
    {
        var option = new Option<OutputFormat>("--format", "-f")
        {
            Description = "Output format: json | text. Defaults to json when stdout is piped, text on a terminal.",
            DefaultValueFactory = _ => OutputFormat.Auto,
            CustomParser = result =>
            {
                if (result.Tokens.Count == 0)
                    return OutputFormat.Auto;

                var value = result.Tokens[0].Value;
                if (value.Equals("json", StringComparison.OrdinalIgnoreCase))
                    return OutputFormat.Json;
                if (value.Equals("text", StringComparison.OrdinalIgnoreCase))
                    return OutputFormat.Text;

                result.AddError($"Invalid value for --format: '{value}'. Allowed values are: json, text.");
                return OutputFormat.Auto;
            },
        };

        return option;
    }

    public static Option<bool> BuildQuietOption() =>
        new("--quiet", "-q")
        {
            Description = "Suppress the success summary on stdout. Errors and exit codes are unaffected.",
        };
}
