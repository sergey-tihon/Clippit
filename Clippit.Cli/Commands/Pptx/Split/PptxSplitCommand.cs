using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Pptx.Split;

/// <summary>
/// <c>clippit pptx split &lt;input.pptx|-&gt; [--output &lt;dir&gt;] [--slides &lt;expr&gt;]
///                       [--force] [--format json|text] [--quiet]</c>
///
/// Splits selected slides in a .pptx file into individual single-slide .pptx files.
/// Pass <c>-</c> as the input to read the source presentation from stdin.
/// </summary>
internal static class PptxSplitCommand
{
    public static Command Build()
    {
        var inputArg = InputSource.BuildArgument(
            "input",
            "Path to the source .pptx file to split. Use '-' to read from stdin."
        );

        var outputOption = new Option<DirectoryInfo?>("--output", "-o")
        {
            Description =
                "Output directory for the individual slide files "
                + "(default: same directory as input, or the current directory when reading from stdin).",
        };

        var slidesOption = new Option<string?>("--slides", "-s")
        {
            Description =
                "Slides to extract. Syntax: comma-separated list of 1-based slide numbers and/or "
                + "inclusive ranges, e.g. '1,3,6-9'. Whitespace allowed around separators. "
                + "Duplicates are deduplicated. Default: all slides.",
        };

        var forceOption = new Option<bool>("--force")
        {
            Description =
                "Overwrite existing output files. Without this flag the command fails on the first collision.",
        };

        var manifestOption = new Option<bool>("--manifest")
        {
            Description =
                "Also write a deck manifest (clippit pptx build run compatible) alongside the slides. "
                + "Path: <outputDir>/<sourceBaseName>.manifest.json (or stdin.manifest.json for stdin input). "
                + "Preserves source PPTX sections when present.",
        };

        var cmd = new Command(
            "split",
            "Split a .pptx file into individual single-slide files."
                + "\n\nExamples:"
                + "\n  clippit pptx split deck.pptx --slides 1,3,6-9 --output slides"
                + "\n  clippit pptx split deck.pptx --output slides --manifest"
                + "\n  cat deck.pptx | clippit pptx split - --output slides --format json"
        );
        cmd.Arguments.Add(inputArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(slidesOption);
        cmd.Options.Add(forceOption);
        cmd.Options.Add(manifestOption);
        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    parseResult.GetValue(inputArg)!,
                    parseResult.GetValue(outputOption),
                    parseResult.GetValue(slidesOption),
                    parseResult.GetValue(forceOption),
                    parseResult.GetValue(manifestOption),
                    parseResult.GetValue(formatOption),
                    parseResult.GetValue(quietOption)
                )
            )
        );

        return cmd;
    }

    private static int Run(
        string inputPath,
        DirectoryInfo? outputOption,
        string? slidesExpression,
        bool force,
        bool writeManifest,
        OutputFormat format,
        bool quiet
    )
    {
        var writer = new OutputWriter(format, quiet);
        var input = InputSource.From(inputPath, "stdin.pptx");
        var outputDir = outputOption ?? ResolveDefaultOutputDir(inputPath, input.IsStdin);
        outputDir.Create();

        var result = PptxSplitService.Execute(input, outputDir, slidesExpression, force, writeManifest);
        writer.WriteResult(result, CliJsonContext.Default.SplitResult, SplitResult.WriteText);
        return ExitCodes.Success;
    }

    private static DirectoryInfo ResolveDefaultOutputDir(string inputPath, bool fromStdin) =>
        fromStdin
            ? new DirectoryInfo(Directory.GetCurrentDirectory())
            : new FileInfo(inputPath).Directory ?? new DirectoryInfo(Directory.GetCurrentDirectory());
}
