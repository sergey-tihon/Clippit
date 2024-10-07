using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint;

public static partial class PresentationBuilder
{
    public static IList<PmlDocument> PublishSlides(PmlDocument src)
    {
        using var streamSrcDoc = new OpenXmlMemoryStreamDocument(src);
        using var srcDoc = streamSrcDoc.GetPresentationDocument(new OpenSettings { AutoSave = false });
        return PublishSlides(srcDoc, src.FileName).ToList();
    }

    public static IEnumerable<PmlDocument> PublishSlides(PresentationDocument srcDoc, string fileName)
    {
        var slideNameRegex = SlideNameRegex();
        var slideNumber = 0;
        foreach (var streamDoc in PublishSlides(srcDoc))
        {
            try
            {
                var slideDoc = streamDoc.GetModifiedPmlDocument();
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    slideDoc.FileName = slideNameRegex.Replace(fileName, $"_{++slideNumber:000}.pptx");
                }

                yield return slideDoc;
            }
            finally
            {
                streamDoc.Dispose();
            }
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
        using var builder = Create(output);
        try
        {
            var newSlidePart = builder.AddSlidePart(slidePart);

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
