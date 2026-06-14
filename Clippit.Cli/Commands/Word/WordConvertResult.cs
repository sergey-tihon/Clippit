namespace Clippit.Cli.Commands.Word;

/// <summary>
/// Result of the "word to-html" or "word from-html" command.
/// Both commands produce the same payload shape: input path, output path, and output file size.
/// </summary>
internal sealed record WordConvertResult
{
    public required string Input { get; init; }
    public required string Output { get; init; }
    public required long OutputSize { get; init; }

    public static void WriteText(WordConvertResult result, TextWriter writer)
    {
        writer.WriteLine($"{result.Input} → {result.Output}");
        writer.WriteLine($"Output size: {result.OutputSize} bytes");
    }
}
