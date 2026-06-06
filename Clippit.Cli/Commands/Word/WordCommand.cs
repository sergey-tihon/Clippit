using System.CommandLine;
using Clippit.Cli.Commands.Word.Verify;

namespace Clippit.Cli.Commands.Word;

/// <summary>
/// Root "word" subcommand group — all Word (.docx) operations live here.
/// </summary>
internal static class WordCommand
{
    public static Command Build()
    {
        var cmd = new Command("word", "Work with Word (.docx) files");
        cmd.Subcommands.Add(WordVerifyCommand.Build());
        return cmd;
    }
}
