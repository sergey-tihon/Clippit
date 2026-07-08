using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.Build;

internal static class WordBuildInitCommand
{
    internal const string DefaultManifestName = "clippit-word-build.json";

    public static Command Build()
    {
        var outputOption = new Option<string?>("--output", "-o")
        {
            Description =
                $"Path for the new manifest file (default: ./{DefaultManifestName}). "
                + "Use '-' to write the manifest to stdout.",
        };

        var forceOption = new Option<bool>("--force")
        {
            Description = "Overwrite the manifest file if it already exists.",
        };

        var cmd = new Command(
            "init",
            "Scaffold a new empty Word build manifest."
                + "\n\nExamples:"
                + "\n  clippit word build init"
                + "\n  clippit word build init --output word-build.json --force"
                + "\n  clippit word build init --output - > word-build.json"
        );
        cmd.Options.Add(outputOption);
        cmd.Options.Add(forceOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Init(
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(forceOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Init(string? outputOption, bool force, OutputFormat format, bool quiet)
    {
        var toStdout = outputOption == OutputTarget.StdoutToken;
        var writer = new OutputWriter(format, quiet || toStdout);
        var result = WordBuildInitService.Execute(outputOption, force);
        writer.WriteResult(result, CliJsonContext.Default.WordBuildInitResult, WordBuildInitResult.WriteText);
        return ExitCodes.Success;
    }
}
