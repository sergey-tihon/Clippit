using System.CommandLine;
using Clippit.Cli.Commands.Pptx.Build;
using Clippit.Cli.Commands.Pptx.Split;
using Clippit.Cli.Commands.Pptx.Verify;

namespace Clippit.Cli.Commands.Pptx;

/// <summary>
/// Root "pptx" subcommand group — all PPTX operations live here.
/// </summary>
internal static class PptxCommand
{
    public static Command Build()
    {
        var cmd = new Command("pptx", "Work with PowerPoint (.pptx) files");
        cmd.Subcommands.Add(PptxSplitCommand.Build());
        cmd.Subcommands.Add(PptxBuildCommand.Build());
        cmd.Subcommands.Add(PptxVerifyCommand.Build());
        return cmd;
    }
}
