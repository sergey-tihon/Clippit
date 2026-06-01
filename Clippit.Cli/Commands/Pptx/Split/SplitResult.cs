namespace Clippit.Cli.Commands.Pptx.Split;

using System.Globalization;

/// <summary>
/// Result of the "pptx split" command.
/// </summary>
internal sealed record SplitResult
{
    public required string Input { get; init; }
    public required string OutputDir { get; init; }

    /// <summary>
    /// Absolute path of the generated deck manifest, when <c>--manifest</c> was requested.
    /// Null otherwise.
    /// </summary>
    public string? Manifest { get; init; }

    public required IReadOnlyList<SlideEntry> Slides { get; init; }
    public required int Count { get; init; }

    public static void WriteText(SplitResult result, TextWriter writer)
    {
        writer.WriteLine($"Split: {result.Input}");
        writer.WriteLine($"Output: {result.OutputDir}");
        if (result.Manifest is not null)
            writer.WriteLine($"Manifest: {result.Manifest}");
        writer.WriteLine($"Slides: {result.Count}");

        var maxIndex = result.Slides.Count > 0 ? result.Slides.Max(s => s.Index) : 0;
        var width = Math.Max(3, maxIndex.ToString(CultureInfo.InvariantCulture).Length);

        foreach (var slide in result.Slides)
        {
            var title = slide.Title is not null ? $" — {slide.Title}" : string.Empty;
            var idx = slide.Index.ToString(CultureInfo.InvariantCulture).PadLeft(width, '0');
            writer.WriteLine($"  [{idx}] {slide.File}{title}");
        }
    }
}

internal sealed record SlideEntry
{
    public required int Index { get; init; }
    public required string File { get; init; }
    public string? Title { get; init; }
}
