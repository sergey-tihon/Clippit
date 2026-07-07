namespace Clippit.Cli.Commands.Word.Consolidate;

internal sealed record RevisionInfoResult
{
    public required string File { get; init; }
    public required string Revisor { get; init; }
    public required string Color { get; init; }
}

internal sealed record ConsolidateResult
{
    public required string Original { get; init; }
    public required IReadOnlyList<RevisionInfoResult> Revisions { get; init; }
    public required string Output { get; init; }
    public required long OutputSize { get; init; }

    public static void WriteText(ConsolidateResult result, TextWriter writer)
    {
        writer.WriteLine($"Original: {result.Original}");
        writer.WriteLine($"Revisions: {result.Revisions.Count}");
        foreach (var rev in result.Revisions)
            writer.WriteLine($"  {rev.File} ({rev.Revisor}, {rev.Color})");
        writer.WriteLine($"Output: {result.Output}");
        writer.WriteLine($"Output size: {result.OutputSize} bytes");
    }
}
