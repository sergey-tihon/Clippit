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

                // Update docProps/app.xml to reflect single-slide metadata
                UpdateExtendedFileProperties(output, title);
            }

            srcSlidePart.RemoveAnnotations<XDocument>();
            srcSlidePart.UnloadRootElement();

            yield return memoryStream;
        }
    }

    [GeneratedRegex(".pptx", RegexOptions.IgnoreCase, "en-US")]
    private static partial Regex SlideNameRegex();

    /// <summary>
    /// Updates docProps/app.xml so that a single-slide output document reflects accurate metadata:
    /// <list type="bullet">
    /// <item>ep:Slides count is set to 1</item>
    /// <item>HeadingPairs "Slide Titles" count is set to 1</item>
    /// <item>TitlesOfParts slide-title entries are replaced with the single slide title</item>
    /// </list>
    /// </summary>
    private static void UpdateExtendedFileProperties(PresentationDocument doc, string slideTitle)
    {
        var extPart = doc.ExtendedFilePropertiesPart;
        if (extPart is null)
            return;

        var xDoc = extPart.GetXDocument();
        var root = xDoc.Root;
        if (root is null)
            return;

        root.Element(EP.Slides)?.SetValue("1");

        var headingPairsVector = root.Elements(EP.HeadingPairs).Elements(VT.vector).FirstOrDefault();
        var titlesVector = root.Elements(EP.TitlesOfParts).Elements(VT.vector).FirstOrDefault();

        if (headingPairsVector is null || titlesVector is null)
        {
            extPart.PutXDocument();
            return;
        }

        var allTitles = titlesVector.Elements(VT.lpstr).Select(e => e.Value).ToList();
        var variants = headingPairsVector.Elements(VT.variant).ToList();

        var newTitles = new List<string>();
        var offset = 0;
        for (var i = 0; i + 1 < variants.Count; i += 2)
        {
            var typeName = variants[i].Element(VT.lpstr)?.Value;
            if (!int.TryParse(variants[i + 1].Element(VT.i4)?.Value, out var count))
                count = 0;

            if (typeName == "Slide Titles")
            {
                newTitles.Add(slideTitle);
                variants[i + 1].Element(VT.i4)?.SetValue("1");
            }
            else
            {
                newTitles.AddRange(allTitles.Skip(offset).Take(count));
            }

            offset += count;
        }

        titlesVector.RemoveNodes();
        titlesVector.SetAttributeValue("size", newTitles.Count.ToString());
        foreach (var t in newTitles)
            titlesVector.Add(new XElement(VT.lpstr, t));

        extPart.PutXDocument();
    }
}
