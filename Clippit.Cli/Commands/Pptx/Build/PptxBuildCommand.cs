using System.CommandLine;

namespace Clippit.Cli.Commands.Pptx.Build;

/// <summary>
/// <c>clippit pptx build init [--output &lt;manifest.json|-&gt;] [--force]</c>
/// <c>clippit pptx build run  &lt;manifest.json|-&gt; [--output &lt;file.pptx|-&gt;] [--format json|text] [--quiet]</c>
/// </summary>
internal static class PptxBuildCommand
{
    public static Command Build()
    {
        var cmd = new Command(
            "build",
            "Build a presentation from a manifest that assembles source .pptx files into sections."
                + "\n\nExamples:"
                + "\n  clippit pptx build init --output deck.json"
                + "\n  clippit pptx build run deck.json --output final.pptx"
        );
        cmd.Subcommands.Add(PptxBuildInitCommand.Build());
        cmd.Subcommands.Add(PptxBuildRunCommand.Build());
        return cmd;
    }
}
