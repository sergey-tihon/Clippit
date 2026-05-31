using System.CommandLine;

namespace Clippit.Cli.Infrastructure;

internal static class CommandExtensions
{
    public static (Option<OutputFormat> Format, Option<bool> Quiet) AddOutputOptions(this Command command)
    {
        var formatOption = FormatOption.BuildFormatOption();
        var quietOption = FormatOption.BuildQuietOption();

        command.Options.Add(formatOption);
        command.Options.Add(quietOption);

        return (formatOption, quietOption);
    }
}
