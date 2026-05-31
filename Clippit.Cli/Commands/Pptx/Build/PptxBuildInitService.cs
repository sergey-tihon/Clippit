using System.Text.Json;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Pptx.Build;

internal static class PptxBuildInitService
{
    public static InitResult Execute(string? outputOption, bool force)
    {
        var target = OutputTarget.FromOption(
            outputOption,
            () => Path.Combine(Directory.GetCurrentDirectory(), PptxBuildInitCommand.DefaultManifestName)
        );

        var manifest = new BuildManifest
        {
            Schema = CliConstants.DeckManifestSchema,
            Title = "My Presentation",
            Output = "final.pptx",
            Deck = [new DeckEntry { Section = "Section 1" }, new DeckEntry { File = "part1.pptx" }],
        };

        var json = JsonSerializer.Serialize(manifest, CliJsonContextIndented.Default.BuildManifest);

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

        return new InitResult { Manifest = target.DisplayPath };
    }
}
