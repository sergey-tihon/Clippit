using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.ToHtml;

internal static class WordToHtmlCommand
{
    public static Command Build()
    {
        var inputArg = InputSource.BuildArgument(
            "input",
            "Path to the .docx file to convert. Use '-' to read from stdin."
        );

        var outputOption = new Option<string>("--output", "-o")
        {
            Description =
                "Output path for the generated .html file. "
                + "Defaults to <input>.html. Use '-' to write HTML content to stdout.",
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
            DefaultValueFactory = _ => "pt-",
        };

        var inlineImagesOption = new Option<bool>("--inline-images")
        {
            Description = "Embed images as base64 data URIs instead of linking to external files.",
        };

        var noFabricateCssOption = new Option<bool>("--no-fabricate-css")
        {
            Description = "Skip CSS class generation and use inline style attributes instead.",
        };

        var cmd = new Command(
            "to-html",
            "Convert a .docx file to HTML/CSS."
                + "\n\nExamples:"
                + "\n  clippit word to-html document.docx"
                + "\n  clippit word to-html document.docx --page-title \"Q3 Report\" --additional-css \"body { max-width: 800px; }\""
                + "\n  clippit word to-html document.docx --inline-images --output -"
                + "\n  cat document.docx | clippit word to-html - --inline-images --format json"
        );
        cmd.Arguments.Add(inputArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(pageTitleOption);
        cmd.Options.Add(additionalCssOption);
        cmd.Options.Add(cssPrefixOption);
        cmd.Options.Add(inlineImagesOption);
        cmd.Options.Add(noFabricateCssOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(inputArg)!,
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(pageTitleOption),
                    parseResult.GetValue(additionalCssOption),
                    parseResult.GetValue(cssPrefixOption) ?? "pt-",
                    !parseResult.GetValue(noFabricateCssOption),
                    parseResult.GetValue(inlineImagesOption),
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
        string? pageTitle,
        string? additionalCss,
        string cssPrefix,
        bool fabricateCss,
        bool inlineImages,
        OutputFormat format,
        bool quiet
    )
    {
        var input = InputSource.From(inputPath, "stdin.docx");
        var defaultOutput = Path.ChangeExtension(input.DisplayName, ".html");
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        var writer = new OutputWriter(format, quiet || output.IsStdout);

        var result = WordToHtmlService.Execute(
            input,
            output,
            pageTitle,
            additionalCss,
            cssPrefix,
            fabricateCss,
            inlineImages
        );
        writer.WriteResult(result, CliJsonContext.Default.WordConvertResult, WordConvertResult.WriteText);
        return ExitCodes.Success;
    }
}
