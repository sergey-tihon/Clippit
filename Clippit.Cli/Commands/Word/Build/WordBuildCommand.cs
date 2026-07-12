using System.CommandLine;

namespace Clippit.Cli.Commands.Word.Build;

/// <summary>
/// <c>clippit word build init [--output &lt;manifest.json|-&gt;] [--force]</c>
/// <c>clippit word build run  &lt;manifest.json|-&gt; [--output &lt;file.docx|-&gt;] [--format json|text] [--quiet]</c>
/// </summary>
internal static class WordBuildCommand
{
    public static Command Build()
    {
        var cmd = new Command(
            "build",
            "Build a merged Word document from a manifest that combines multiple .docx sources."
                + "\n\nExamples:"
                + "\n  clippit word build init --output word-build.json"
                + "\n  clippit word build run word-build.json --output merged.docx"
        );
        cmd.Subcommands.Add(WordBuildInitCommand.Build());
        cmd.Subcommands.Add(WordBuildRunCommand.Build());
        return cmd;
    }
}
