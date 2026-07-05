using System.CommandLine;
using Clippit.Cli.Commands.Word.AcceptRevisions;
using Clippit.Cli.Commands.Word.Compare;
using Clippit.Cli.Commands.Word.FromHtml;
using Clippit.Cli.Commands.Word.ToHtml;
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
        cmd.Subcommands.Add(WordAcceptRevisionsCommand.Build());
        cmd.Subcommands.Add(WordCompareCommand.Build());
        cmd.Subcommands.Add(WordVerifyCommand.Build());
        cmd.Subcommands.Add(WordToHtmlCommand.Build());
        cmd.Subcommands.Add(WordFromHtmlCommand.Build());
        return cmd;
    }
}
