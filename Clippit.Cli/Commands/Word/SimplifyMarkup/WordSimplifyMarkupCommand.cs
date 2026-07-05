using System.CommandLine;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.SimplifyMarkup;

internal static class WordSimplifyMarkupCommand
{
    public static Command Build()
    {
        var inputArg = InputSource.BuildArgument("input", "Path to the .docx file. Use '-' to read from stdin.");

        var outputOption = new Option<string?>("--output", "-o")
        {
            Description =
                "Output path for the simplified .docx. "
                + "Defaults to <input>-simplified.docx. Use '-' to write binary content to stdout.",
        };

        var forceOption = new Option<bool>("--force")
        {
            Description = "Overwrite the output file if it already exists.",
        };

        // Simplification flags
        var allOption = new Option<bool>("--all") { Description = "Enable all cleanup flags (convenience preset)." };
        var acceptRevisionsOption = new Option<bool>("--accept-revisions")
        {
            Description = "Accept all tracked revisions before simplification.",
        };
        var removeRsidInfoOption = new Option<bool>("--remove-rsid-info")
        {
            Description = "Remove RSID attributes from settings and content.",
        };
        var removeMarkupForDocCompOption = new Option<bool>("--remove-markup-for-document-comparison")
        {
            Description = "Remove document properties and markup used for comparison (implies --remove-rsid-info).",
        };
        var removeCommentsOption = new Option<bool>("--remove-comments")
        {
            Description = "Remove comments and comment-extended markup.",
        };
        var removeBookmarksOption = new Option<bool>("--remove-bookmarks")
        {
            Description = "Remove bookmarks (including _GoBack).",
        };
        var removeContentControlsOption = new Option<bool>("--remove-content-controls")
        {
            Description = "Remove structured document tags, keeping their content.",
        };
        var removeEndAndFootnotesOption = new Option<bool>("--remove-end-and-footnotes")
        {
            Description = "Remove endnotes and footnotes.",
        };
        var removeFieldCodesOption = new Option<bool>("--remove-field-codes")
        {
            Description = "Remove field codes, leaving the last cached result text.",
        };
        var removeGoBackBookmarkOption = new Option<bool>("--remove-go-back-bookmark")
        {
            Description = "Remove the _GoBack bookmark specifically.",
        };
        var removeHyperlinksOption = new Option<bool>("--remove-hyperlinks")
        {
            Description = "Remove hyperlink relationships and markup.",
        };
        var removeLastRenderedPageBreakOption = new Option<bool>("--remove-last-rendered-page-break")
        {
            Description = "Remove lastRenderedPageBreak elements.",
        };
        var removePermissionsOption = new Option<bool>("--remove-permissions")
        {
            Description = "Remove permission and editable-region markup.",
        };
        var removeProofOption = new Option<bool>("--remove-proof") { Description = "Remove proofing errors." };
        var removeSmartTagsOption = new Option<bool>("--remove-smart-tags")
        {
            Description = "Remove smart-tag wrappers.",
        };
        var removeSoftHyphensOption = new Option<bool>("--remove-soft-hyphens")
        {
            Description = "Remove soft hyphen characters.",
        };
        var removeWebHiddenOption = new Option<bool>("--remove-web-hidden") { Description = "Remove web-hidden text." };
        var replaceTabsWithSpacesOption = new Option<bool>("--replace-tabs-with-spaces")
        {
            Description = "Replace tab characters with spaces.",
        };
        var normalizeXmlOption = new Option<bool>("--normalize-xml")
        {
            Description = "Apply the library's XML normalization step.",
        };

        var cmd = new Command(
            "simplify-markup",
            "Simplify markup in a .docx file by removing non-content elements."
                + "\n\nAt least one simplification flag must be provided, or use --all."
                + "\n\nExamples:"
                + "\n  clippit word simplify-markup noisy.docx --remove-rsid-info --remove-comments --output clean.docx"
                + "\n  clippit word simplify-markup noisy.docx --all --output clean.docx"
                + "\n  clippit word simplify-markup noisy.docx --all --format json"
        );

        cmd.Arguments.Add(inputArg);
        cmd.Options.Add(outputOption);
        cmd.Options.Add(forceOption);
        cmd.Options.Add(allOption);
        cmd.Options.Add(acceptRevisionsOption);
        cmd.Options.Add(removeRsidInfoOption);
        cmd.Options.Add(removeMarkupForDocCompOption);
        cmd.Options.Add(removeCommentsOption);
        cmd.Options.Add(removeBookmarksOption);
        cmd.Options.Add(removeContentControlsOption);
        cmd.Options.Add(removeEndAndFootnotesOption);
        cmd.Options.Add(removeFieldCodesOption);
        cmd.Options.Add(removeGoBackBookmarkOption);
        cmd.Options.Add(removeHyperlinksOption);
        cmd.Options.Add(removeLastRenderedPageBreakOption);
        cmd.Options.Add(removePermissionsOption);
        cmd.Options.Add(removeProofOption);
        cmd.Options.Add(removeSmartTagsOption);
        cmd.Options.Add(removeSoftHyphensOption);
        cmd.Options.Add(removeWebHiddenOption);
        cmd.Options.Add(replaceTabsWithSpacesOption);
        cmd.Options.Add(normalizeXmlOption);

        var (formatOption, quietOption) = cmd.AddOutputOptions();

        cmd.SetAction(parseResult =>
            CommandRunner.Execute(() =>
                Run(
                    new WordSimplifyMarkupOptions(
                        parseResult.GetValue(inputArg)!,
                        parseResult.GetValue(outputOption),
                        parseResult.GetValue(forceOption),
                        parseResult.GetValue(allOption),
                        parseResult.GetValue(acceptRevisionsOption),
                        parseResult.GetValue(removeRsidInfoOption),
                        parseResult.GetValue(removeMarkupForDocCompOption),
                        parseResult.GetValue(removeCommentsOption),
                        parseResult.GetValue(removeBookmarksOption),
                        parseResult.GetValue(removeContentControlsOption),
                        parseResult.GetValue(removeEndAndFootnotesOption),
                        parseResult.GetValue(removeFieldCodesOption),
                        parseResult.GetValue(removeGoBackBookmarkOption),
                        parseResult.GetValue(removeHyperlinksOption),
                        parseResult.GetValue(removeLastRenderedPageBreakOption),
                        parseResult.GetValue(removePermissionsOption),
                        parseResult.GetValue(removeProofOption),
                        parseResult.GetValue(removeSmartTagsOption),
                        parseResult.GetValue(removeSoftHyphensOption),
                        parseResult.GetValue(removeWebHiddenOption),
                        parseResult.GetValue(replaceTabsWithSpacesOption),
                        parseResult.GetValue(normalizeXmlOption),
                        parseResult.GetValue(formatOption),
                        parseResult.GetValue(quietOption)
                    )
                )
            )
        );

        return cmd;
    }

    private static int Run(WordSimplifyMarkupOptions options)
    {
        var input = InputSource.From(options.InputPath, "stdin.docx");
        var defaultOutput = input.IsStdin
            ? "simplified.docx"
            : Path.Combine(
                Path.GetDirectoryName(input.DisplayName)!,
                $"{Path.GetFileNameWithoutExtension(input.DisplayName)}-simplified.docx"
            );
        var output = OutputTarget.FromOption(options.OutputPath, () => defaultOutput);
        var writer = new OutputWriter(options.Format, options.Quiet || output.IsStdout);

        var result = WordSimplifyMarkupService.Execute(input, output, options);

        writer.WriteResult(result, CliJsonContext.Default.ConvertResult, ConvertResult.WriteText);
        return ExitCodes.Success;
    }
}
