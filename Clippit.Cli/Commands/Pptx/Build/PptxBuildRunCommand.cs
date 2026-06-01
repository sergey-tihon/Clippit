using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Pptx.Build;

/// <summary>
/// <c>clippit pptx build run &lt;manifest.json|-&gt; [--output &lt;file.pptx|-&gt;] [--format json|text] [--quiet]</c>
///
/// Builds a final presentation from a deck manifest. Pass <c>-</c> as the manifest
/// path to read it from stdin; pass <c>--output -</c> to write the resulting
/// .pptx to stdout. Binary output to stdout suppresses the success summary.
/// </summary>
internal static class PptxBuildRunCommand
{
    public static Command Build()
    {
        var manifestArg = InputSource.BuildArgument(
            "manifest",
            "Path to the deck manifest JSON file. Use '-' to read the manifest from stdin.",
            "Manifest"
        );

        var outputOption = new Option<string?>("--output", "-o")
        {
            Description = "Override the output .pptx path from the manifest. Use '-' to write the binary to stdout.",
        };

        var forceOption = new Option<bool>("--force")
        {
            Description = "Overwrite the output presentation if it already exists.",
        };

        var cmd = new Command(
            "run",
            "Build the final presentation from a deck manifest."
                + "\n\nExamples:"
                + "\n  clippit pptx build run deck.json"
                + "\n  clippit pptx build run deck.json --output final.pptx --quiet"
                + "\n  cat deck.json | clippit pptx build run - --output - > final.pptx"
        );
        cmd.Arguments.Add(manifestArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(forceOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(manifestArg)!,
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(forceOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Run(string manifestPath, string? outputOverride, bool force, OutputFormat format, bool quiet)
    {
        var manifestInput = InputSource.From(manifestPath, "stdin.json");
        var toStdout = outputOverride == OutputTarget.StdoutToken;
        var writer = new OutputWriter(format, quiet || toStdout);

        var buildResult = PptxBuildRunService.Execute(manifestInput, outputOverride, force);
        writer.WriteResult(buildResult, CliJsonContext.Default.BuildResult, BuildResult.WriteText);
        return ExitCodes.Success;
    }
}
