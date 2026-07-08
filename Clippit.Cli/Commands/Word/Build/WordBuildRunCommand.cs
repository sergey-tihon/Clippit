using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.Build;

/// <summary>
/// <c>clippit word build run &lt;manifest.json|-&gt; [--output &lt;file.docx|-&gt;] [--format json|text] [--quiet]</c>
///
/// Builds a merged Word document from a manifest. Pass <c>-</c> as the manifest
/// path to read it from stdin; pass <c>--output -</c> to write the resulting
/// .docx to stdout. Binary output to stdout suppresses the success summary.
/// </summary>
internal static class WordBuildRunCommand
{
    public static Command Build()
    {
        var manifestArg = InputSource.BuildArgument(
            "manifest",
            "Path to the Word build manifest JSON file. Use '-' to read the manifest from stdin.",
            "Manifest"
        );

        var outputOption = new Option<string?>("--output", "-o")
        {
            Description = "Override the output .docx path from the manifest. Use '-' to write the binary to stdout.",
        };

        var forceOption = new Option<bool>("--force")
        {
            Description = "Overwrite the output document if it already exists.",
        };

        var cmd = new Command(
            "run",
            "Build the merged Word document from a manifest."
                + "\n\nExamples:"
                + "\n  clippit word build run word-build.json"
                + "\n  clippit word build run word-build.json --output merged.docx --quiet"
                + "\n  cat word-build.json | clippit word build run - --output - > merged.docx"
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

        var buildResult = WordBuildRunService.Execute(manifestInput, outputOverride, force);
        writer.WriteResult(buildResult, CliJsonContext.Default.WordBuildResult, WordBuildResult.WriteText);
        return ExitCodes.Success;
    }
}
