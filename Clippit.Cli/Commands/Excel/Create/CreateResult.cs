namespace Clippit.Cli.Commands.Excel.Create;

internal sealed record CreateResult
{
    public required string Input { get; init; }
    public required string Output { get; init; }
    public required long OutputSize { get; init; }
    public required int WorksheetCount { get; init; }

    public static void WriteText(CreateResult result, TextWriter writer)
    {
        writer.WriteLine($"Input: {result.Input}");
        writer.WriteLine($"Output: {result.Output}");
        writer.WriteLine($"Output size: {result.OutputSize} bytes");
        writer.WriteLine($"Worksheets: {result.WorksheetCount}");
    }
}
