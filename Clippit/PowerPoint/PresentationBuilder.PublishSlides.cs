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
        foreach (var slideId in PresentationBuilderTools.GetSlideIdsInOrder(srcDoc))
        {
            var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slideId);
            var slideName = slideNameRegex.Replace(fileName, $"_{++slideNumber:000}.pptx");

            var memoryStream = new MemoryStream();
            try
            {
                PublishSlideToStream(srcSlidePart, memoryStream);
                srcSlidePart.RemoveAnnotations<XDocument>();
                srcSlidePart.UnloadRootElement();
                yield return new PmlDocument(slideName, memoryStream);
            }
            finally
            {
                memoryStream.Dispose();
            }
        }
    }

    /// <summary>
    /// Extracts every slide from a presentation file directly to individual .pptx files on disk.
    /// </summary>
    /// <remarks>
    /// <para>
    /// This overload is optimised for large presentations (1 GB+) and slides containing large
    /// embedded media (videos, audio).  Unlike <see cref="PublishSlides(PmlDocument)"/>, which
    /// requires the entire source document to be loaded into a byte array, this method opens the
    /// source file directly from disk via <see cref="PresentationDocument.Open(string, bool)"/>
    /// and writes each output slide directly to a <see cref="FileStream"/>, so peak memory usage
    /// stays proportional to one slide rather than the whole presentation.
    /// </para>
    /// </remarks>
    /// <param name="srcPath">Path to the source .pptx file.</param>
    /// <param name="outputDirectory">
    /// Directory where the per-slide files will be written.  Created if it does not exist.
    /// </param>
    /// <returns>
    /// The full path of each written slide file, yielded in slide order.  The iterator is lazy:
    /// each slide file is written on demand as the caller advances the enumerator, so you can
    /// stop early without processing the whole presentation.
    /// </returns>
    public static IEnumerable<string> PublishSlides(string srcPath, string outputDirectory)
    {
        ArgumentNullException.ThrowIfNull(srcPath);
        ArgumentNullException.ThrowIfNull(outputDirectory);
        Directory.CreateDirectory(outputDirectory);

        // Open directly from disk — avoids loading the entire file into a MemoryStream.
        using var srcDoc = PresentationDocument.Open(srcPath, isEditable: false);
        var fileName = Path.GetFileName(srcPath);
        foreach (var outputPath in PublishSlides(srcDoc, fileName, outputDirectory))
            yield return outputPath;
    }

    /// <summary>
    /// Extracts every slide from an already-open presentation to individual .pptx files on disk.
    /// </summary>
    /// <remarks>
    /// Each slide is written directly to a <see cref="FileStream"/>, so slides with large
    /// embedded media (1 GB+ videos) do not require an equivalently large heap allocation.
    /// </remarks>
    /// <param name="srcDoc">The open source presentation.</param>
    /// <param name="fileName">
    /// Base file name used to derive per-slide names (e.g. <c>deck.pptx</c> → <c>deck_001.pptx</c>).
    /// </param>
    /// <param name="outputDirectory">
    /// Directory where the per-slide files will be written.  Created if it does not exist.
    /// </param>
    /// <returns>
    /// The full path of each written slide file, yielded in slide order.
    /// </returns>
    public static IEnumerable<string> PublishSlides(
        PresentationDocument srcDoc,
        string fileName,
        string outputDirectory
    )
    {
        ArgumentNullException.ThrowIfNull(srcDoc);
        ArgumentNullException.ThrowIfNull(outputDirectory);
        Directory.CreateDirectory(outputDirectory);

        fileName ??= string.Empty;
        var slideNameRegex = SlideNameRegex();
        var slideNumber = 0;

        foreach (var slideId in PresentationBuilderTools.GetSlideIdsInOrder(srcDoc))
        {
            var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slideId);
            var slideName = slideNameRegex.Replace(fileName, $"_{++slideNumber:000}.pptx");
            var outputPath = Path.Combine(outputDirectory, slideName);

            // Write directly to a FileStream — no large MemoryStream on the heap.
            using (var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                PublishSlideToStream(srcSlidePart, fileStream);

            srcSlidePart.RemoveAnnotations<XDocument>();
            srcSlidePart.UnloadRootElement();

            yield return outputPath;
        }
    }

    /// <summary>Builds a single-slide presentation from <paramref name="srcSlidePart"/> and writes it to <paramref name="outputStream"/>.</summary>
    private static void PublishSlideToStream(SlidePart srcSlidePart, Stream outputStream)
    {
        using var output = NewDocument(outputStream);
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
        titlesVector.SetAttributeValue(NoNamespace.size, newTitles.Count);
        foreach (var t in newTitles)
            titlesVector.Add(new XElement(VT.lpstr, t));

        extPart.PutXDocument();
    }
}
