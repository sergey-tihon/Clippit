using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
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
        fileName ??= string.Empty;

        var slideNameRegex = SlideNameRegex();
        var slideNumber = 0;
        foreach (var memoryStream in PublishSlides(srcDoc))
        {
            try
            {
                var slideName = slideNameRegex.Replace(fileName, $"_{++slideNumber:000}.pptx");
                yield return new PmlDocument(slideName, memoryStream);
            }
            finally
            {
                memoryStream.Dispose();
            }
        }
    }

    private static IEnumerable<MemoryStream> PublishSlides(PresentationDocument srcDoc)
    {
        var slidesIds = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc);
        foreach (var slideId in slidesIds)
        {
            var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slideId);

            var memoryStream = new MemoryStream();
            using (var output = NewDocument(memoryStream))
            {
                using (var builder = Create(output))
                {
                    var newSlidePart = builder.AddSlidePart(srcSlidePart);

                    // Remove the show attribute from the slide element (if it exists)
                    var slideDocument = newSlidePart.GetXDocument();
                    slideDocument.Root?.Attribute(NoNamespace.show)?.Remove();
                }

                // Set the title of the new presentation to the title of the slide
                var title = PresentationBuilderTools.GetSlideTitle(srcSlidePart.GetXElement());
                output.PackageProperties.Title = title;
            }

            srcSlidePart.RemoveAnnotations<XDocument>();
            srcSlidePart.UnloadRootElement();

            yield return memoryStream;
        }
    }

    [GeneratedRegex(".pptx", RegexOptions.IgnoreCase, "en-US")]
    private static partial Regex SlideNameRegex();
}
