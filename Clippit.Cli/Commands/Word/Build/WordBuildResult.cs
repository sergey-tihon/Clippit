namespace Clippit.Cli.Commands.Word.Build;

/// <summary>Result of the "word build run" command.</summary>
internal sealed record WordBuildResult
{
    public required string Output { get; init; }
    public required long OutputSize { get; init; }
    public required int EntryCount { get; init; }
    public required IReadOnlyList<WordBuildEntryResult> Entries { get; init; }

    public static void WriteText(WordBuildResult result, TextWriter writer)
    {
        writer.WriteLine($"Build → {result.Output}");
        writer.WriteLine($"Output size: {result.OutputSize:N0} bytes");
        if (result.Entries.Count == 0)
            return;

        writer.WriteLine("Deck:");
        foreach (var entry in result.Entries)
        {
            if (entry.Section is not null)
                writer.WriteLine($"  [section] {entry.Section}");
            else
                writer.WriteLine(
                    $"  [file]    {entry.File}  ({entry.Elements} element{(entry.Elements == 1 ? "" : "s")})"
                );
        }
    }
}

internal sealed record WordBuildEntryResult
{
    public string? Section { get; init; }
    public string? File { get; init; }
    public int? Elements { get; init; }
}

/// <summary>Result of the "word build init" command.</summary>
internal sealed record WordBuildInitResult
{
    public required string Manifest { get; init; }

    public static void WriteText(WordBuildInitResult result, TextWriter writer) =>
        writer.WriteLine($"Created manifest: {result.Manifest}");
}
