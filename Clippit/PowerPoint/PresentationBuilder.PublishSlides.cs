using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Clippit.Internal;
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

            // Normalize all relationship IDs in the output ZIP so that IDs are stable
            // (deterministic) across runs. The OpenXML SDK generates GUID-based IDs for
            // parts added via AddNewPart<T>(), which would otherwise vary each invocation.
            var normalizedStream = NormalizeRelationshipIds(memoryStream);
            memoryStream.Dispose();
            yield return normalizedStream;
        }
    }

    private static readonly XNamespace s_packageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    private static readonly DateTimeOffset s_zipEpoch = new(1980, 1, 1, 0, 0, 0, TimeSpan.Zero);

    /// <summary>
    /// Rewrites the OOXML ZIP package so that every relationship ID is a stable, deterministic
    /// value derived from SHA-256 of the relationship type and target URI, rather than the
    /// random GUID the OpenXML SDK assigns when creating parts with AddNewPart&lt;T&gt;().
    /// Entry timestamps are also normalised to the ZIP epoch so the output is byte-for-byte
    /// identical across invocations with the same input.
    /// </summary>
    private static MemoryStream NormalizeRelationshipIds(MemoryStream input)
    {
        input.Position = 0;

        // Pass 1: read every .rels file and build a mapping scoped by .rels entry path.
        // Keys are (relsEntryPath, oldId) so that identical old IDs like "rId1" in different
        // .rels files remain independent and don't overwrite each other.
        // Also build a lookup from .rels path → the set of XML entry paths it governs.
        var scopedMappings = new Dictionary<string, Dictionary<string, string>>();

        using (var inZip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
        {
            foreach (var entry in inZip.Entries.Where(e => e.FullName.EndsWith(".rels")))
            {
                using var stream = entry.Open();
                var doc = XDocument.Load(stream);
                if (doc.Root is null)
                    continue;

                // Assign rId1, rId2, … in document order within each .rels file.
                var gen = new RelationshipIdGenerator();
                var mappings = new Dictionary<string, string>();
                foreach (var rel in doc.Root.Elements(s_packageRelNs + "Relationship"))
                {
                    var oldId = (string)rel.Attribute("Id")!;
                    mappings[oldId] = gen.Next();
                }

                scopedMappings[entry.FullName] = mappings;
            }
        }

        // Build a map from each XML/rels entry → the set of .rels files whose mappings apply.
        // A .rels file at "_rels/.rels" governs root-level entries; a .rels file at
        // "ppt/slides/_rels/slide1.xml.rels" governs "ppt/slides/slide1.xml", etc.
        // Each .rels always governs itself, plus the owner part it describes.
        var entryToMappings = new Dictionary<string, Dictionary<string, string>>(StringComparer.Ordinal);

        foreach (var (relsPath, mappings) in scopedMappings)
        {
            // The .rels file itself needs rewriting
            AddMappingsForEntry(entryToMappings, relsPath, mappings);

            // Derive the owner part path:
            // "_rels/.rels" → "" (package-level)
            // "ppt/_rels/presentation.xml.rels" → "ppt/presentation.xml"
            // "ppt/slides/_rels/slide1.xml.rels" → "ppt/slides/slide1.xml"
            var relsDir = relsPath.Replace("_rels/", "");
            var ownerPath = relsDir.EndsWith(".rels") ? relsDir[..^".rels".Length] : relsDir;
            if (ownerPath.Length > 0)
            {
                AddMappingsForEntry(entryToMappings, ownerPath, mappings);
            }
        }

        // Pass 2: rewrite the ZIP with updated IDs and normalised entry timestamps.
        var output = new MemoryStream();
        input.Position = 0;

        using (var inZip = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true))
        using (var outZip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            // Process entries in a stable order so the ZIP layout is deterministic.
            foreach (var entry in inZip.Entries.OrderBy(e => e.FullName, StringComparer.Ordinal))
            {
                var newEntry = outZip.CreateEntry(entry.FullName);
                newEntry.LastWriteTime = s_zipEpoch;

                using var inStream = entry.Open();
                using var outStream = newEntry.Open();

                if (
                    (entry.FullName.EndsWith(".rels") || entry.FullName.EndsWith(".xml"))
                    && entryToMappings.TryGetValue(entry.FullName, out var applicableMappings)
                )
                {
                    var text = new StreamReader(inStream, Encoding.UTF8).ReadToEnd();
                    foreach (var (oldId, newId) in applicableMappings)
                    {
                        text = text.Replace($"=\"{oldId}\"", $"=\"{newId}\"");
                    }
                    outStream.Write(Encoding.UTF8.GetBytes(text));
                }
                else
                {
                    inStream.CopyTo(outStream);
                }
            }
        }

        output.Position = 0;
        return output;
    }

    private static void AddMappingsForEntry(
        Dictionary<string, Dictionary<string, string>> entryToMappings,
        string entryPath,
        Dictionary<string, string> mappings
    )
    {
        if (!entryToMappings.TryGetValue(entryPath, out var existing))
        {
            existing = new Dictionary<string, string>();
            entryToMappings[entryPath] = existing;
        }

        foreach (var (key, value) in mappings)
        {
            existing[key] = value;
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

        var headingPairsVector = root.Element(EP.HeadingPairs)?.Element(VT.vector);
        var titlesVector = root.Element(EP.TitlesOfParts)?.Element(VT.vector);

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
