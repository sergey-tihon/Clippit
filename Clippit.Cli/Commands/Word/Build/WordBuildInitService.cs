using System.Text.Json;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Word.Build;

internal static class WordBuildInitService
{
    public static WordBuildInitResult Execute(string? outputOption, bool force)
    {
        var target = OutputTarget.FromOption(
            outputOption,
            () => Path.Combine(Directory.GetCurrentDirectory(), WordBuildInitCommand.DefaultManifestName)
        );

        var manifest = new WordBuildManifest
        {
            Schema = CliConstants.WordBuildManifestSchema,
            Output = "merged.docx",
            Entries = [new WordEntryItem { Section = "Part 1" }, new WordEntryItem { File = "part1.docx" }],
        };

        var json = JsonSerializer.Serialize(manifest, CliJsonContextIndented.Default.WordBuildManifest);

        if (target.IsStdout)
        {
            Console.Out.WriteLine(json);
        }
        else
        {
            target.EnsureCanWrite(force, "Manifest");
            target.EnsureDirectoryExists();
            File.WriteAllText(target.DisplayPath, json);
        }

        return new WordBuildInitResult { Manifest = target.DisplayPath };
    }
}
