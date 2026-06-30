using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.Compare;

internal static class WordCompareCommand
{
    public static Command Build()
    {
        var sourceArg = InputSource.BuildArgument(
            "source",
            "Path to the source .docx file. Use '-' to read from stdin."
        );
        var revisedArg = InputSource.BuildArgument(
            "revised",
            "Path to the revised .docx file. Use '-' to read from stdin."
        );

        var outputOption = new Option<string>("--output", "-o")
        {
            Description =
                "Output path for the compared .docx file with tracked revisions. "
                + "Defaults to <source>-compared.docx. Use '-' to write binary content to stdout.",
        };

        var authorOption = new Option<string>("--author")
        {
            Description = "Author value used for generated tracked revisions.",
        };

        var dateTimeOption = new Option<string>("--date-time")
        {
            Description = "Date/time value used for generated tracked revisions (ISO 8601 text recommended).",
        };

        var caseInsensitiveOption = new Option<bool>("--case-insensitive")
        {
            Description = "Ignore case when comparing words.",
        };

        var cmd = new Command(
            "compare",
            "Compare two .docx files and produce a result with tracked revisions."
                + "\n\nExamples:"
                + "\n  clippit word compare before.docx after.docx"
                + "\n  clippit word compare before.docx after.docx --output compared.docx --format json"
                + "\n  clippit word compare before.docx after.docx --author \"Jane Doe\" --date-time 2026-01-01T00:00:00Z"
                + "\n  cat before.docx | clippit word compare - after.docx --output compared.docx --format json"
        );
        cmd.Arguments.Add(sourceArg);
        cmd.Arguments.Add(revisedArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(authorOption);
        cmd.Options.Add(dateTimeOption);
        cmd.Options.Add(caseInsensitiveOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(sourceArg)!,
                    parseResult.GetValue(revisedArg)!,
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(authorOption),
                    parseResult.GetValue(dateTimeOption),
                    parseResult.GetValue(caseInsensitiveOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Run(
        string sourcePath,
        string revisedPath,
        string? outputPath,
        string? authorForRevisions,
        string? dateTimeForRevisions,
        bool caseInsensitive,
        OutputFormat format,
        bool quiet
    )
    {
        var source = InputSource.From(sourcePath, "stdin-source.docx");
        var revised = InputSource.From(revisedPath, "stdin-revised.docx");
        var defaultOutput = source.IsStdin
            ? "compared.docx"
            : Path.Combine(
                Path.GetDirectoryName(source.DisplayName)!,
                $"{Path.GetFileNameWithoutExtension(source.DisplayName)}-compared.docx"
            );
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        var writer = new OutputWriter(format, quiet || output.IsStdout);

        var result = WordCompareService.Execute(
            source,
            revised,
            output,
            authorForRevisions,
            dateTimeForRevisions,
            caseInsensitive
        );
        writer.WriteResult(result, CliJsonContext.Default.CompareResult, CompareResult.WriteText);
        return ExitCodes.Success;
    }
}
