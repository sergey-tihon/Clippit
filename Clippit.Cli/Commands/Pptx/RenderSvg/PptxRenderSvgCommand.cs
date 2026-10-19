using System.CommandLine;
using Clippit.Cli.Infrastructure;
using Clippit.PowerPoint;

namespace Clippit.Cli.Commands.Pptx.RenderSvg;

internal static class PptxRenderSvgCommand
{
    public static Command Build()
    {
        var inputArg = new Argument<string[]>("input")
        {
            Description = "One or more SVG files to render in order. Use '-' for a single SVG from stdin.",
            Arity = ArgumentArity.OneOrMore,
        };
        inputArg.Validators.Add(result =>
        {
            var values = result.GetValueOrDefault<string[]>() ?? [];
            if (values.Count(value => value == InputSource.StdinToken) > 1)
                result.AddError("Only one SVG input may be read from stdin.");

            foreach (var value in values)
            {
                if (value != InputSource.StdinToken && !File.Exists(value))
                    result.AddError($"SVG file not found: {value}");
            }
        });
        var outputOption = new Option<string>("--output", "-o")
        {
            Description = "Output path for the generated .pptx file. Defaults to the input name with .pptx extension.",
        };
        var forceOption = new Option<bool>("--force", "-f") { Description = "Overwrite an existing output file." };
        var templateOption = new Option<string?>("--template")
        {
            Description = "Existing .pptx whose selected slides receive the generated SVG shapes.",
        };
        var templateSlideOption = new Option<int>("--template-slide")
        {
            Description = "Zero-based template slide index to start overlaying (default: 0).",
            DefaultValueFactory = _ => 0,
        };
        templateSlideOption.Validators.Add(result =>
        {
            if (result.GetValueOrDefault<int>() < 0)
                result.AddError("--template-slide must be zero or greater.");
        });
        templateOption.Validators.Add(result =>
        {
            var value = result.GetValueOrDefault<string?>();
            if (value is not null && !File.Exists(value))
                result.AddError($"Template file not found: {value}");
        });
        var cmd = new Command(
            "render-svg",
            "Render a standalone SVG as editable PowerPoint shapes."
                + "\n\nExamples:"
                + "\n  clippit pptx render-svg architecture.svg -o architecture.pptx"
                + "\n  clippit pptx render-svg title.svg architecture.svg -o deck.pptx"
                + "\n  cat architecture.svg | clippit pptx render-svg - -o architecture.pptx"
        );
        cmd.Arguments.Add(inputArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(forceOption);
        cmd.Options.Add(templateOption);
        cmd.Options.Add(templateSlideOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();
        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(inputArg) ?? [],
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(templateOption),
                    parseResult.GetValue(templateSlideOption),
                    parseResult.GetValue(forceOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );
        return cmd;
    }

    private static int Run(
        string[] inputPaths,
        string? outputPath,
        string? templatePath,
        int templateSlideIndex,
        bool force,
        OutputFormat format,
        bool quiet
    )
    {
        if (inputPaths.Length == 0)
            throw CliException.InvalidArguments("At least one SVG input must be supplied.");

        var inputs = inputPaths.Select(path => InputSource.From(path, "stdin.svg")).ToArray();
        var template = templatePath is null ? null : new PmlDocument(Path.GetFullPath(templatePath));
        var defaultOutput =
            inputs.Length == 1 && !inputs[0].IsStdin
                ? Path.ChangeExtension(inputs[0].DisplayName, ".pptx")
                : Path.Combine(Directory.GetCurrentDirectory(), "rendered.pptx");
        var output = OutputTarget.FromOption(outputPath, () => defaultOutput);
        output.EnsureDirectoryExists();
        output.EnsureCanWrite(force, "Output file");
        var result = PptxRenderSvgService.Execute(inputs, output, template, templateSlideIndex);
        var writer = new OutputWriter(format, quiet || output.IsStdout);
        writer.WriteResult(result, CliJsonContext.Default.RenderSvgResult, RenderSvgResult.WriteText);
        return ExitCodes.Success;
    }
}
