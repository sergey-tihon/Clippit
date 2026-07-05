using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.AcceptRevisions;

internal static class WordAcceptRevisionsCommand
{
    public static Command Build()
    {
        var inputArg = InputSource.BuildArgument("input", "Path to the .docx file. Use '-' to read from stdin.");

        var outputOption = new Option<string?>("--output", "-o")
        {
            Description =
                "Output path for the cleaned .docx. "
                + "Defaults to <input>-accepted.docx. Use '-' to write binary content to stdout.",
        };

        var forceOption = new Option<bool>("--force")
        {
            Description = "Overwrite the output file if it already exists.",
        };

        var cmd = new Command(
            "accept-revisions",
            "Accept all tracked revisions in a .docx file."
                + "\n\nExamples:"
                + "\n  clippit word accept-revisions draft.docx"
                + "\n  clippit word accept-revisions draft.docx --output clean.docx --format json"
                + "\n  cat draft.docx | clippit word accept-revisions - --output clean.docx"
                + "\n  clippit word accept-revisions draft.docx --output - > clean.docx"
        );
        cmd.Arguments.Add(inputArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(forceOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(inputArg)!,
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(forceOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Run(string inputPath, string? outputPath, bool force, OutputFormat format, bool quiet)
    {
        var input = InputSource.From(inputPath, "stdin.docx");
        var defaultOutput = input.IsStdin
            ? "accepted.docx"
            : Path.Combine(
                Path.GetDirectoryName(input.DisplayName)!,
                $"{Path.GetFileNameWithoutExtension(input.DisplayName)}-accepted.docx"
            );
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        var writer = new OutputWriter(format, quiet || output.IsStdout);

        var result = WordAcceptRevisionsService.Execute(input, output, force);
        writer.WriteResult(result, CliJsonContext.Default.ConvertResult, ConvertResult.WriteText);
        return ExitCodes.Success;
    }
}
