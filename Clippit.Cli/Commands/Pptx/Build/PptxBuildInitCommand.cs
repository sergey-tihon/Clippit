using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Pptx.Build;

internal static class PptxBuildInitCommand
{
    internal const string DefaultManifestName = "clippit-deck.json";

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
            "Scaffold a new empty deck manifest."
                + "\n\nExamples:"
                + "\n  clippit pptx build init"
                + "\n  clippit pptx build init --output deck.json --force"
                + "\n  clippit pptx build init --output - > deck.json"
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
        var result = PptxBuildInitService.Execute(outputOption, force);
        writer.WriteResult(result, CliJsonContext.Default.InitResult, InitResult.WriteText);
        return ExitCodes.Success;
    }
}
