using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.FromHtml;

internal static class WordFromHtmlCommand
{
    public static Command Build()
    {
        var inputArg = InputSource.BuildArgument(
            "input",
            "Path to the .html file to convert. Use '-' to read from stdin."
        );

        var outputOption = new Option<string>("--output", "-o")
        {
            Description =
                "Output path for the generated .docx file. "
                + "Defaults to <input>.docx. Use '-' to write binary content to stdout.",
        };

        var cssOption = new Option<string>("--css", "-c")
        {
            Description =
                "Path to an external author CSS file. "
                + "When omitted, CSS is extracted from the HTML <style> element.",
        };

        var defaultCssOption = new Option<string>("--default-css")
        {
            Description = "Path to a default CSS file. " + "When omitted, a built-in default CSS is used.",
        };

        var userCssOption = new Option<string>("--user-css")
        {
            Description = "Additional CSS rules to apply as user overrides.",
        };

        var baseUriOption = new Option<string>("--base-uri")
        {
            Description =
                "Base URI for resolving relative image src references. "
                + "Defaults to the source HTML file's parent directory.",
        };

        var majorFontOption = new Option<string>("--major-font")
        {
            Description = "Theme major (heading) font name (default: Calibri Light).",
        };

        var minorFontOption = new Option<string>("--minor-font")
        {
            Description = "Theme minor (body) font name (default: Times New Roman).",
        };

        var fontSizeOption = new Option<double?>("--font-size")
        {
            Description = "Default font size in points (default: 12).",
        };

        var cmd = new Command(
            "from-html",
            "Convert an HTML file to .docx."
                + "\n\nExamples:"
                + "\n  clippit word from-html article.html"
                + "\n  clippit word from-html article.html -c styles.css -o article.docx"
                + "\n  clippit word from-html article.html --minor-font \"Georgia\" --font-size 11"
                + "\n  cat article.html | clippit word from-html - --base-uri https://example.com/images/ --format json"
        );
        cmd.Arguments.Add(inputArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(cssOption);
        cmd.Options.Add(defaultCssOption);
        cmd.Options.Add(userCssOption);
        cmd.Options.Add(baseUriOption);
        cmd.Options.Add(majorFontOption);
        cmd.Options.Add(minorFontOption);
        cmd.Options.Add(fontSizeOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(inputArg)!,
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(cssOption),
                    parseResult.GetValue(defaultCssOption),
                    parseResult.GetValue(userCssOption),
                    parseResult.GetValue(baseUriOption),
                    parseResult.GetValue(majorFontOption),
                    parseResult.GetValue(minorFontOption),
                    parseResult.GetValue(fontSizeOption),
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
        string? cssFilePath,
        string? defaultCssFilePath,
        string? userCss,
        string? baseUri,
        string? majorFont,
        string? minorFont,
        double? fontSize,
        OutputFormat format,
        bool quiet
    )
    {
        var input = InputSource.From(inputPath, "stdin.html");
        var defaultOutput = Path.ChangeExtension(input.IsStdin ? input.LogicalName : input.DisplayName, ".docx");
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        var writer = new OutputWriter(format, quiet || output.IsStdout);

        // When no --base-uri is given and input is a file, default to the file's directory
        if (baseUri is null && !input.IsStdin)
        {
            var fileInfo = new FileInfo(inputPath);
            if (fileInfo.DirectoryName is not null)
                baseUri = fileInfo.DirectoryName;
        }

        var result = WordFromHtmlService.Execute(
            input,
            output,
            cssFilePath,
            defaultCssFilePath,
            userCss,
            baseUri,
            majorFont,
            minorFont,
            fontSize
        );

        writer.WriteResult(result, CliJsonContext.Default.ConvertResult, ConvertResult.WriteText);
        return ExitCodes.Success;
    }
}
