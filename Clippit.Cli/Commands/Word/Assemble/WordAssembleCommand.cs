using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.Assemble;

internal static class WordAssembleCommand
{
    public static Command Build()
    {
        var templateArg = InputSource.BuildArgument(
            "template",
            "Path to the template .docx file. Use '-' to read from stdin."
        );
        var dataArg = InputSource.BuildArgument(
            "data",
            "Path to the XML data file. Use '-' to read from stdin.",
            "XML data file"
        );

        var outputOption = new Option<string>("--output", "-o")
        {
            Description =
                "Output path for the assembled .docx file. "
                + "Defaults to <template>-assembled.docx. Use '-' to write binary content to stdout.",
        };

        var forceOption = new Option<bool>("--force", "-f")
        {
            Description = "Overwrite the output file if it already exists.",
        };

        var cmd = new Command(
            "assemble",
            "Assemble a .docx document from a template and XML data."
                + "\n\nExamples:"
                + "\n  clippit word assemble template.docx data.xml"
                + "\n  clippit word assemble template.docx data.xml --output generated.docx --format json"
                + "\n  clippit word assemble template.docx - --output generated.docx"
                + "\n  clippit word assemble template.docx data.xml --output - > generated.docx"
        );
        cmd.Arguments.Add(templateArg);
        cmd.Arguments.Add(dataArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(forceOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(templateArg)!,
                    parseResult.GetValue(dataArg)!,
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(forceOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Run(
        string templatePath,
        string dataPath,
        string? outputPath,
        bool force,
        OutputFormat format,
        bool quiet
    )
    {
        var template = InputSource.From(templatePath, "stdin-template.docx");
        var data = InputSource.From(dataPath, "stdin-data.xml");
        var defaultOutput = template.IsStdin
            ? "assembled.docx"
            : Path.Combine(
                Path.GetDirectoryName(template.DisplayName)!,
                $"{Path.GetFileNameWithoutExtension(template.DisplayName)}-assembled.docx"
            );
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        var writer = new OutputWriter(format, quiet || output.IsStdout);

        var result = WordAssembleService.Execute(template, data, output, force);
        writer.WriteResult(result, CliJsonContext.Default.AssembleResult, AssembleResult.WriteText);
        return ExitCodes.Success;
    }
}
