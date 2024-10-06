using System.Collections.Generic;
using System.Text.RegularExpressions;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint;

public static partial class PresentationBuilder
{
    public static IEnumerable<PmlDocument> PublishSlides(PmlDocument src)
    {
        using var streamSrcDoc = new OpenXmlMemoryStreamDocument(src);
        using var srcDoc = streamSrcDoc.GetPresentationDocument(new OpenSettings { AutoSave = false });
        return PublishSlides(srcDoc, src.FileName);
    }

    public static IEnumerable<PmlDocument> PublishSlides(PresentationDocument srcDoc, string fileName)
    {
        var slideNameRegex = SlideNameRegex();
        var slideNumber = 0;
        foreach (var memoryStreamDocument in PublishSlides(srcDoc))
        {
            using var streamDoc = memoryStreamDocument;

            var slideDoc = streamDoc.GetModifiedPmlDocument();
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                slideDoc.FileName = slideNameRegex.Replace(fileName, $"_{++slideNumber:000}.pptx");
            }

            yield return slideDoc;
        }
    }

    public static IEnumerable<OpenXmlMemoryStreamDocument> PublishSlides(PresentationDocument srcDoc)
    {
        var slidesIds = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc);
        foreach (var slideId in slidesIds)
        {
            var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slideId);

            var streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();

            using (var output = streamDoc.GetPresentationDocument(new OpenSettings { AutoSave = false }))
            {
                ExtractSlide(srcSlidePart, output);

                // Set the title of the new presentation to the title of the slide
                var title = PresentationBuilderTools.GetSlideTitle(srcSlidePart.GetXElement());
                output.PackageProperties.Title = title;
            }

            yield return streamDoc;
        }
    }

    private static void ExtractSlide(SlidePart slidePart, PresentationDocument output)
    {
        using var builder = new FluentPresentationBuilder(output);
        try
        {
            var newSlidePart = builder.AddSlide(slidePart);

            // Remove the show attribute from the slide element (if it exists)
            var slideDocument = newSlidePart.GetXDocument();
            slideDocument.Root?.Attribute(NoNamespace.show)?.Remove();
        }
        catch (PresentationBuilderInternalException dbie)
        {
            if (dbie.Message.Contains("{0}"))
                throw new PresentationBuilderException(string.Format(dbie.Message, slidePart.Uri));
            throw;
        }
    }

    [GeneratedRegex(".pptx", RegexOptions.IgnoreCase, "en-US")]
    private static partial Regex SlideNameRegex();
}
