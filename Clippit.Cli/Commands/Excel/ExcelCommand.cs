using System.CommandLine;
using Clippit.Cli.Commands.Excel.Verify;

namespace Clippit.Cli.Commands.Excel;

/// <summary>
/// Root "excel" subcommand group — all Excel (.xlsx) operations live here.
/// </summary>
internal static class ExcelCommand
{
    public static Command Build()
    {
        var cmd = new Command("excel", "Work with Excel (.xlsx) files");
        cmd.Subcommands.Add(ExcelVerifyCommand.Build());
        return cmd;
    }
}
