using System.CommandLine;
using System.Text.Json;
using Clippit.Cli;
using Clippit.Cli.Commands.Excel;
using Clippit.Cli.Commands.Pptx;
using Clippit.Cli.Commands.Version;
using Clippit.Cli.Commands.Word;

var rootCommand = new RootCommand("Clippit — PowerTools CLI for OpenXml (PowerPoint, Word, Excel)");

rootCommand.Subcommands.Add(PptxCommand.Build());
rootCommand.Subcommands.Add(WordCommand.Build());
rootCommand.Subcommands.Add(ExcelCommand.Build());
rootCommand.Subcommands.Add(VersionCommand.Build());

args = NormalizeHelpArgs(args);

// Top-level --version: intercept to ensure the payload is identical to
// `clippit version` (structured JSON). System.CommandLine 3.x adds a built-in
// VersionOption that would otherwise print a plain version string and diverge
// from the subcommand.
if (IsTopLevelVersionRequest(args))
{
    Console.WriteLine(JsonSerializer.Serialize(VersionCommand.BuildResult(), CliJsonContext.Default.VersionResult));
    return 0;
}

return rootCommand.Parse(args).Invoke();

// Map `help` and `help <cmd...>` to the conventional `--help` form so users
// (and LLMs) can use either style.
static string[] NormalizeHelpArgs(string[] arguments)
{
    if (arguments is not ["help", .. var rest])
        return arguments;
    return rest.Length == 0 ? ["--help"] : [.. rest, "--help"];
}

static bool IsTopLevelVersionRequest(string[] arguments)
{
    // Only intercept when the user really meant the root --version, i.e. no
    // subcommand precedes it and no help follows.
    return arguments.Length > 0
        && arguments[0] is "--version"
        && !arguments.Any(argument => argument is "--help" or "-h" or "-?");
}

// Trick to give the implicit top-level class a stable name for AOT
internal sealed partial class Program { }
