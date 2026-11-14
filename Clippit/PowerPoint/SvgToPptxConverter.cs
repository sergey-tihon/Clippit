using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;
using Drawing = Clippit.A;
using Presentation = Clippit.P;
using Relationships = Clippit.R;

namespace Clippit.PowerPoint;

/// <summary>Converts a supported, standalone SVG into an editable PresentationML presentation.</summary>
public static partial class SvgToPptxConverter
{
    // PresentationML/DrawingML/relationship names come from PtOpenXmlUtil; SVG is a separate namespace.
    private static readonly XNamespace S = "http://www.w3.org/2000/svg";
    private const long DefaultSlideWidth = 12_192_000;
    private const long DefaultSlideHeight = 6_858_000;

    public static PmlDocument Convert(string svg, SvgToPptxConverterSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(svg);
        return Convert([new SvgSlideSource(svg)], settings);
    }

    public static PmlDocument Convert(Stream svg, SvgToPptxConverterSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(svg);
        using var reader = new StreamReader(svg, leaveOpen: true);
        return Convert(reader.ReadToEnd(), settings);
    }

    /// <summary>
    /// Converts one or more SVG slide sources. When a template is supplied, each SVG is overlaid on a consecutive
    /// template slide starting at <see cref="SvgToPptxConverterSettings.TemplateSlideIndex" />.
    /// </summary>
    public static PmlDocument Convert(IEnumerable<SvgSlideSource> slides, SvgToPptxConverterSettings? settings = null)
    {
        ArgumentNullException.ThrowIfNull(slides);
        settings ??= new SvgToPptxConverterSettings();
        var svgSources = slides.ToList();
        if (svgSources.Count == 0)
            throw new ArgumentException("At least one SVG slide is required.", nameof(slides));

        PmlDocument result;
        if (settings.Template is null)
        {
            using var streamDocument = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
            using (var presentation = streamDocument.GetPresentationDocument(new OpenSettings { AutoSave = false }))
            {
                AddSlidesToNewPresentation(presentation, svgSources, settings);
            }

            result = streamDocument.GetModifiedPmlDocument();
        }
        else
        {
            using var streamDocument = new OpenXmlMemoryStreamDocument(settings.Template);
            using (var presentation = streamDocument.GetPresentationDocument(new OpenSettings { AutoSave = false }))
            {
                OverlaySlidesOnTemplate(presentation, svgSources, settings);
            }

            result = streamDocument.GetModifiedPmlDocument();
        }
        if (settings.ValidateOutput)
        {
            var validation = PresentationValidator.Validate(new MemoryStream(result.DocumentByteArray));
            if (!validation.Valid)
            {
                var message = string.Join(Environment.NewLine, validation.Diagnostics.Select(d => d.Description));
                throw new InvalidDataException($"Generated presentation is invalid:{Environment.NewLine}{message}");
            }
        }

        return result;
    }

    private static void AddSlidesToNewPresentation(
        PresentationDocument presentation,
        IReadOnlyList<SvgSlideSource> svgSources,
        SvgToPptxConverterSettings settings
    )
    {
        var presentationPart =
            presentation.PresentationPart ?? throw new InvalidOperationException("Presentation part is missing.");
        CreatePresentationParts(presentationPart);
        var presentationXml = presentationPart.GetXDocument();
        var slideMasterPart = presentationPart.SlideMasterParts.Single();
        var slideLayoutPart = slideMasterPart.SlideLayoutParts.Single();
        var slideIdList = presentationXml.Root!.Element(Presentation.sldIdLst)!;
        uint slideId = 256;
        uint shapeId = 2;

        foreach (var source in svgSources)
        {
            var document = SvgParser.Parse(source.Content, source.Name, settings);
            var slidePart = presentationPart.AddNewPart<SlidePart>();
            slidePart.AddPart(slideLayoutPart);
            var slideXml = CreateSlideXml();
            var writer = new SlideWriter(document, shapeId, DefaultSlideWidth, DefaultSlideHeight);
            writer.Write(slideXml.Root!.Element(Presentation.cSld)!.Element(Presentation.spTree)!);
            shapeId = writer.NextShapeId;
            slidePart.PutXDocument(slideXml);
            slideIdList.Add(
                new XElement(
                    Presentation.sldId,
                    new XAttribute(NoNamespace.id, slideId++),
                    new XAttribute(Relationships.id, presentationPart.GetIdOfPart(slidePart))
                )
            );
        }

        presentationXml.Root!.Element(Presentation.sldSz)?.SetAttributeValue(NoNamespace.cx, DefaultSlideWidth);
        presentationXml.Root!.Element(Presentation.sldSz)?.SetAttributeValue(NoNamespace.cy, DefaultSlideHeight);
        presentationPart.PutXDocument(presentationXml);
    }

    private static void OverlaySlidesOnTemplate(
        PresentationDocument presentation,
        IReadOnlyList<SvgSlideSource> svgSources,
        SvgToPptxConverterSettings settings
    )
    {
        var presentationPart =
            presentation.PresentationPart
            ?? throw new InvalidOperationException("Template presentation part is missing.");
        var slides = GetSlidesInPresentationOrder(presentationPart);
        if (settings.TemplateSlideIndex < 0 || settings.TemplateSlideIndex >= slides.Count)
            throw new ArgumentOutOfRangeException(
                nameof(settings.TemplateSlideIndex),
                settings.TemplateSlideIndex,
                $"Template contains {slides.Count} slide(s)."
            );

        var availableSlides = slides.Count - settings.TemplateSlideIndex;
        if (svgSources.Count > availableSlides)
            throw new ArgumentException(
                $"Template has {availableSlides} slide(s) available from index {settings.TemplateSlideIndex}, "
                    + $"but {svgSources.Count} SVG slide(s) were supplied.",
                nameof(svgSources)
            );

        var (slideWidth, slideHeight) = GetSlideSize(presentationPart);
        for (var index = 0; index < svgSources.Count; index++)
        {
            var slidePart = slides[settings.TemplateSlideIndex + index];
            var slideXml = slidePart.GetXDocument();
            var shapeTree = slideXml.Root?.Element(Presentation.cSld)?.Element(Presentation.spTree);
            if (shapeTree is null)
                throw new InvalidDataException("Template slide does not contain a shape tree.");

            var document = SvgParser.Parse(svgSources[index].Content, svgSources[index].Name, settings);
            var writer = new SlideWriter(
                document,
                PresentationBuilderTools.GetNextShapeId(shapeTree),
                slideWidth,
                slideHeight
            );
            writer.Write(shapeTree);
            slidePart.PutXDocument(slideXml);
        }
    }

    private static List<SlidePart> GetSlidesInPresentationOrder(PresentationPart presentationPart)
    {
        var presentation = presentationPart.GetXDocument().Root;
        var slideIdList = presentation?.Element(Presentation.sldIdLst)?.Elements(Presentation.sldId) ?? [];
        return slideIdList
            .Select(slideId =>
                presentationPart.GetPartById((string)slideId.Attribute(Relationships.id)!) as SlidePart
                ?? throw new InvalidDataException("Presentation contains a non-slide relationship.")
            )
            .ToList();
    }

    private static (long Width, long Height) GetSlideSize(PresentationPart presentationPart)
    {
        var size = presentationPart.GetXDocument().Root?.Element(Presentation.sldSz);
        if (size is null)
            return (DefaultSlideWidth, DefaultSlideHeight);

        return (
            ParsePositiveLong(size.Attribute(NoNamespace.cx)?.Value, DefaultSlideWidth),
            ParsePositiveLong(size.Attribute(NoNamespace.cy)?.Value, DefaultSlideHeight)
        );
    }

    private static long ParsePositiveLong(string? value, long fallback) =>
        long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) && parsed > 0
            ? parsed
            : fallback;
}
