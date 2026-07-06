namespace Clippit.Cli.Commands.Word.Compare;

internal sealed record CompareResult
{
    public required string Source { get; init; }
    public required string Revised { get; init; }
    public required string Output { get; init; }
    public required long OutputSize { get; init; }
    public required int Revisions { get; init; }
    public required string AuthorForRevisions { get; init; }
    public required string DateTimeForRevisions { get; init; }
    public required bool CaseInsensitive { get; init; }

    public static void WriteText(CompareResult result, TextWriter writer)
    {
        writer.WriteLine($"{result.Source} vs {result.Revised}");
        writer.WriteLine($"Output: {result.Output}");
        writer.WriteLine($"Output size: {result.OutputSize} bytes");
        writer.WriteLine($"Revisions: {result.Revisions}");
        writer.WriteLine($"Author: {result.AuthorForRevisions}");
        writer.WriteLine($"Date/time: {result.DateTimeForRevisions}");
        writer.WriteLine($"Case-insensitive: {result.CaseInsensitive}");
    }
}
