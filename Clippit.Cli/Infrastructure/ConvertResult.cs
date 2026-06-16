namespace Clippit.Cli.Infrastructure;

/// <summary>
/// Result payload of any convert command (e.g. <c>word to-html</c>, <c>word from-html</c>,
/// <c>excel to-html</c>). All commands produce the same shape: input path, output path,
/// and output file size.
/// </summary>
internal sealed record ConvertResult
{
    public required string Input { get; init; }
    public required string Output { get; init; }
    public required long OutputSize { get; init; }

    public static void WriteText(ConvertResult result, TextWriter writer)
    {
        writer.WriteLine($"{result.Input} → {result.Output}");
        writer.WriteLine($"Output size: {result.OutputSize} bytes");
    }
}
