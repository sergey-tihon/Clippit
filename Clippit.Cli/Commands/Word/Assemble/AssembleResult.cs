namespace Clippit.Cli.Commands.Word.Assemble;

internal sealed record AssembleResult
{
    public required string Template { get; init; }
    public required string Data { get; init; }
    public required string Output { get; init; }
    public required long OutputSize { get; init; }
    public required bool TemplateError { get; init; }

    public static void WriteText(AssembleResult result, TextWriter writer)
    {
        writer.WriteLine($"Created: {result.Output}");
        writer.WriteLine($"Output size: {result.OutputSize} bytes");
        if (result.TemplateError)
            writer.WriteLine("Warning: template contained markup errors (see templateError in JSON mode)");
    }
}
