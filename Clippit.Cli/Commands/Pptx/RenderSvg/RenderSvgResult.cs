namespace Clippit.Cli.Commands.Pptx.RenderSvg;

internal sealed record RenderSvgResult
{
    public required string Input { get; init; }
    public required string Output { get; init; }

    /// <summary>Number of SVG sources rendered or template slides overlaid.</summary>
    public required int SlideCount { get; init; }

    public static void WriteText(RenderSvgResult result, TextWriter writer)
    {
        writer.WriteLine($"Rendered: {result.Input}");
        writer.WriteLine($"Output: {result.Output}");
        writer.WriteLine($"Slides: {result.SlideCount}");
    }
}
