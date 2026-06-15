using System.CommandLine;
using Clippit.Cli.Commands.Word;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Excel.ToHtml;

internal static class ExcelToHtmlCommand
{
    public static Command Build()
    {
        var inputArg = InputSource.BuildArgument(
            "input",
            "Path to the .xlsx file to convert. Use '-' to read from stdin."
        );

        var outputOption = new Option<string>("--output", "-o")
        {
            Description =
                "Output path for the generated .html file. Defaults to <input>.html. Use '-' to write HTML content to stdout.",
        };

        var sheetOption = new Option<string?>("--sheet")
        {
            Description =
                "The sheet name to convert. Defaults to the first sheet found if neither --range nor --table are specified.",
        };

        var rangeOption = new Option<string?>("--range")
        {
            Description = "Coordinates of the cell/range to convert (e.g. A1:D10). Requires --sheet option.",
        };

        var tableOption = new Option<string?>("--table")
        {
            Description = "Name of a defined Excel table to convert. Cannot be combined with --sheet or --range.",
        };

        var pageTitleOption = new Option<string>("--page-title")
        {
            Description = "HTML page <title>. Defaults to the source file name.",
        };

        var additionalCssOption = new Option<string>("--additional-css")
        {
            Description = "Extra CSS rules injected into the generated <style> block.",
        };

        var cssPrefixOption = new Option<string>("--css-prefix")
        {
            Description = "Prefix for auto-generated CSS class names (default: pt-).",
        };

        var noFabricateCssOption = new Option<bool>("--no-fabricate-css")
        {
            Description = "Skip CSS class generation and use inline style attributes instead.",
        };

        var cmd = new Command(
            "to-html",
            "Convert a spreadsheet range, sheet, or table to HTML/CSS."
                + "\n\nExamples:"
                + "\n  clippit excel to-html spreadsheet.xlsx -o sheet.html"
                + "\n  clippit excel to-html spreadsheet.xlsx --sheet \"Q3 Data\" --range \"B2:F15\""
                + "\n  clippit excel to-html spreadsheet.xlsx --table \"SalesTable\" -o - > table.html"
        );
        cmd.Arguments.Add(inputArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(sheetOption);
        cmd.Options.Add(rangeOption);
        cmd.Options.Add(tableOption);
        cmd.Options.Add(pageTitleOption);
        cmd.Options.Add(additionalCssOption);
        cmd.Options.Add(cssPrefixOption);
        cmd.Options.Add(noFabricateCssOption);

        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(inputArg)!,
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(sheetOption),
                    parseResult.GetValue(rangeOption),
                    parseResult.GetValue(tableOption),
                    parseResult.GetValue(pageTitleOption),
                    parseResult.GetValue(additionalCssOption),
                    parseResult.GetValue(cssPrefixOption) ?? "pt-",
                    !parseResult.GetValue(noFabricateCssOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Run(
        string inputPath,
        string? outputPath,
        string? sheetName,
        string? range,
        string? tableName,
        string? pageTitle,
        string? additionalCss,
        string cssPrefix,
        bool fabricateCss,
        OutputFormat format,
        bool quiet
    )
    {
        var input = InputSource.From(inputPath, "stdin.xlsx");
        var defaultOutput = Path.ChangeExtension(input.IsStdin ? input.LogicalName : input.DisplayName, ".html");
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        var writer = new OutputWriter(format, quiet || output.IsStdout);

        var result = ExcelToHtmlService.Execute(
            input,
            output,
            sheetName,
            range,
            tableName,
            pageTitle,
            additionalCss,
            cssPrefix,
            fabricateCss
        );
        writer.WriteResult(result, CliJsonContext.Default.WordConvertResult, WordConvertResult.WriteText);
        return ExitCodes.Success;
    }
}
