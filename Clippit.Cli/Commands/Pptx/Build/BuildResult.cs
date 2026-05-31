namespace Clippit.Cli.Commands.Pptx.Build;

/// <summary>
/// Result of the "pptx build run" command.
/// </summary>
internal sealed record BuildResult
{
    public required string Output { get; init; }
    public required int TotalSlides { get; init; }

    /// <summary>
    /// Per-entry breakdown in deck order. Section entries report their name;
    /// file entries report the manifest-relative path and the slide count
    /// copied from that file.
    /// </summary>
    public required IReadOnlyList<BuildEntryResult> Entries { get; init; }

    public static void WriteText(BuildResult result, TextWriter writer)
    {
        writer.WriteLine($"Build → {result.Output}");
        writer.WriteLine($"Total slides: {result.TotalSlides}");
        if (result.Entries.Count == 0)
            return;

        writer.WriteLine("Deck:");
        foreach (var entry in result.Entries)
        {
            if (entry.Section is not null)
                writer.WriteLine($"  [section] {entry.Section}");
            else
                writer.WriteLine($"  [file]    {entry.File}  ({entry.Slides} slide{(entry.Slides == 1 ? "" : "s")})");
        }
    }
}

internal sealed record BuildEntryResult
{
    public string? Section { get; init; }
    public string? File { get; init; }
    public int Slides { get; init; }
}

/// <summary>
/// Result of the "pptx build init" command.
/// </summary>
internal sealed record InitResult
{
    public required string Manifest { get; init; }

    public static void WriteText(InitResult result, TextWriter writer) =>
        writer.WriteLine($"Created manifest: {result.Manifest}");
}
