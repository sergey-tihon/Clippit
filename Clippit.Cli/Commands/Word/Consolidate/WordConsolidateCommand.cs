using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.Consolidate;

internal static class WordConsolidateCommand
{
    public static Command Build()
    {
        var originalArg = InputSource.BuildArgument(
            "original",
            "Path to the original .docx file. Use '-' to read from stdin."
        );
        var revisionsArg = new Argument<string[]>("revisions") { Description = "One or more revision .docx files." };
        revisionsArg.Validators.Add(result =>
        {
            var values = result.GetValueOrDefault<string[]>() ?? [];
            foreach (var value in values)
            {
                if (value != InputSource.StdinToken && !File.Exists(value))
                    result.AddError($"Revision file not found: {value}");
            }
        });

        var outputOption = new Option<string?>("--output", "-o")
        {
            Description =
                "Output path for the consolidated .docx file. "
                + "Defaults to <original>-consolidated.docx. Use '-' to write binary content to stdout.",
        };

        var forceOption = new Option<bool>("--force")
        {
            Description = "Overwrite the output file if it already exists.",
        };

        var revisorOption = new Option<string[]?>("--revisor")
        {
            Description =
                "Reviewer name for the corresponding revision file. May be repeated; defaults to the revision file name.",
            AllowMultipleArgumentsPerToken = false,
        };
        revisorOption.Arity = ArgumentArity.ZeroOrMore;

        var colorOption = new Option<string[]?>("--color")
        {
            Description =
                "Hex color (#RRGGBB) for the corresponding revision. May be repeated; defaults to a rotating palette.",
            AllowMultipleArgumentsPerToken = false,
        };
        colorOption.Arity = ArgumentArity.ZeroOrMore;

        var authorOption = new Option<string?>("--author")
        {
            Description = "Author value used for generated tracked revisions.",
        };

        var dateTimeOption = new Option<string?>("--date-time")
        {
            Description = "Date/time value used for generated tracked revisions (ISO 8601 text recommended).",
        };

        var caseInsensitiveOption = new Option<bool>("--case-insensitive")
        {
            Description = "Ignore case when comparing words.",
        };

        var noTableConsolidationOption = new Option<bool>("--no-table-consolidation")
        {
            Description = "Disable table-based consolidation layout.",
        };

        var cmd = new Command(
            "consolidate",
            "Combine multiple revisions of a document into one file with tracked changes."
                + "\n\nExamples:"
                + "\n  clippit word consolidate original.docx alice.docx bob.docx"
                + "\n  clippit word consolidate original.docx alice.docx bob.docx --output consolidated.docx --format json"
                + "\n  clippit word consolidate original.docx alice.docx bob.docx --revisor Alice --revisor Bob --color '#FF0000' --color '#0000FF'"
                + "\n  cat original.docx | clippit word consolidate - alice.docx bob.docx --output consolidated.docx"
        );
        cmd.Arguments.Add(originalArg);
        cmd.Arguments.Add(revisionsArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(forceOption);
        cmd.Options.Add(revisorOption);
        cmd.Options.Add(colorOption);
        cmd.Options.Add(authorOption);
        cmd.Options.Add(dateTimeOption);
        cmd.Options.Add(caseInsensitiveOption);
        cmd.Options.Add(noTableConsolidationOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(originalArg)!,
                    parseResult.GetValue(revisionsArg) ?? [],
                    parseResult.GetValue(revisorOption) ?? [],
                    parseResult.GetValue(colorOption) ?? [],
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(forceOption),
                    parseResult.GetValue(authorOption),
                    parseResult.GetValue(dateTimeOption),
                    parseResult.GetValue(caseInsensitiveOption),
                    parseResult.GetValue(noTableConsolidationOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Run(
        string originalPath,
        string[] revisionPaths,
        string[] revisors,
        string[] colors,
        string? outputPath,
        bool force,
        string? authorForRevisions,
        string? dateTimeForRevisions,
        bool caseInsensitive,
        bool noTableConsolidation,
        OutputFormat format,
        bool quiet
    )
    {
        if (revisionPaths.Length == 0)
            throw CliException.InvalidArguments("At least one revision document must be supplied.");

        var original = InputSource.From(originalPath, "stdin-original.docx");

        var revisions = revisionPaths
            .Select((path, i) => InputSource.From(path, $"stdin-revision-{i + 1}.docx"))
            .ToList();

        var defaultOutput = original.IsStdin
            ? "consolidated.docx"
            : Path.Combine(
                Path.GetDirectoryName(original.DisplayName)!,
                $"{Path.GetFileNameWithoutExtension(original.DisplayName)}-consolidated.docx"
            );
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        var writer = new OutputWriter(format, quiet || output.IsStdout);

        var result = WordConsolidateService.Execute(
            original,
            revisions,
            revisors,
            colors,
            output,
            force,
            authorForRevisions,
            dateTimeForRevisions,
            caseInsensitive,
            noTableConsolidation
        );
        writer.WriteResult(result, CliJsonContext.Default.ConsolidateResult, ConsolidateResult.WriteText);
        return ExitCodes.Success;
    }
}
