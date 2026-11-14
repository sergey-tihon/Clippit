using Clippit.Cli.Infrastructure;
using Clippit.PowerPoint;

namespace Clippit.Cli.Commands.Pptx.RenderSvg;

internal static class PptxRenderSvgService
{
    public static RenderSvgResult Execute(
        IReadOnlyList<InputSource> inputs,
        OutputTarget output,
        PmlDocument? template = null,
        int templateSlideIndex = 0
    )
    {
        ArgumentNullException.ThrowIfNull(inputs);
        if (inputs.Count == 0)
            throw CliException.InvalidArguments("At least one SVG input must be supplied.");

        var sources = new List<SvgSlideSource>(inputs.Count);
        foreach (var input in inputs)
        {
            using var source = input.OpenSeekable();
            using var reader = new StreamReader(source);
            sources.Add(new SvgSlideSource(reader.ReadToEnd(), input.LogicalName));
        }

        var document = SvgToPptxConverter.Convert(
            sources,
            new SvgToPptxConverterSettings { Template = template, TemplateSlideIndex = templateSlideIndex }
        );
        string? tempPath = null;
        try
        {
            using (var outputStream = output.OpenWrite(out tempPath))
            {
                outputStream.Write(document.DocumentByteArray);
                outputStream.Flush();
                if (output.IsStdout)
                    output.Flush(outputStream);
            }
            output.Commit(tempPath);
            tempPath = null;
        }
        catch
        {
            OutputTarget.DeleteTemp(tempPath);
            throw;
        }

        return new RenderSvgResult
        {
            Input = string.Join(", ", inputs.Select(input => input.DisplayName)),
            Output = output.DisplayPath,
            SlideCount = inputs.Count,
        };
    }
}
