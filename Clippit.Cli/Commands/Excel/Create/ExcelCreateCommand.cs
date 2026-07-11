using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Excel.Create;

internal static class ExcelCreateCommand
{
    public static Command Build()
    {
        var inputArg = InputSource.BuildArgument(
            "input",
            "Path to the workbook definition JSON file. Use '-' to read from stdin."
        );

        var outputOption = new Option<string?>("--output", "-o")
        {
            Description =
                "Output path for the generated .xlsx file. "
                + "Defaults to the input file name with '.xlsx' extension. "
                + "Use '-' to write binary content to stdout.",
        };

        var forceOption = new Option<bool>("--force")
        {
            Description = "Overwrite the output file if it already exists.",
        };

        var cmd = new Command(
            "create",
            "Generate an Excel (.xlsx) workbook from a JSON workbook definition."
                + "\n\nExamples:"
                + "\n  clippit excel create report.json"
                + "\n  clippit excel create report.json --output report.xlsx --format json"
                + "\n  cat report.json | clippit excel create - --output report.xlsx"
                + "\n  clippit excel create report.json --output - > report.xlsx"
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
        var input = InputSource.From(inputPath, "stdin-workbook.json");
        var defaultOutput = input.IsStdin
            ? "workbook.xlsx"
            : Path.Combine(
                Path.GetDirectoryName(input.DisplayName)!,
                Path.GetFileNameWithoutExtension(input.DisplayName) + ".xlsx"
            );
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        var writer = new OutputWriter(format, quiet || output.IsStdout);

        var result = ExcelCreateService.Execute(input, output, force);
        writer.WriteResult(result, CliJsonContext.Default.CreateResult, CreateResult.WriteText);
        return ExitCodes.Success;
    }
}
