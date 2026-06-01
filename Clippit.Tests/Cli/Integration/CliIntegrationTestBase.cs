using System.IO.Compression;
using System.Text.Json;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Cli.Integration;

internal abstract class CliIntegrationTestBase : TestsBase
{
    protected static async Task<FileInfo> WriteManifestAsync(DirectoryInfo directory, string output, FileInfo source)
    {
        var manifest = new FileInfo(Path.Combine(directory.FullName, "deck.json"));
        var json = JsonSerializer.Serialize(
            new
            {
                title = "CLI Test Presentation",
                output,
                deck = new[] { "[CLI Section]", source.FullName },
            }
        );
        await File.WriteAllTextAsync(manifest.FullName, json).ConfigureAwait(false);
        return manifest;
    }

    protected static void InjectDanglingOleObjRelId(FileInfo pptx, params string[] danglingIds)
    {
        XNamespace pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var zip = ZipFile.Open(pptx.FullName, ZipArchiveMode.Update);
        var slideEntry = zip.Entries.First(e =>
            e.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
            && e.FullName.EndsWith(".xml", StringComparison.Ordinal)
        );

        XDocument xDoc;
        using (var stream = slideEntry.Open())
            xDoc = XDocument.Load(stream);

        var target = xDoc.Descendants(pNs + "spTree").FirstOrDefault() ?? xDoc.Root;
        foreach (var danglingId in danglingIds)
            target?.Add(new XElement(pNs + "oleObj", new XAttribute(rNs + "id", danglingId)));

        var fullName = slideEntry.FullName;
        slideEntry.Delete();
        var newEntry = zip.CreateEntry(fullName);
        using var writer = new StreamWriter(newEntry.Open());
        using var xmlWriter = System.Xml.XmlWriter.Create(writer);
        xDoc.WriteTo(xmlWriter);
    }

    protected static bool HasSectionList(FileInfo pptx)
    {
        using var document = PresentationDocument.Open(pptx.FullName, false);
        var xDoc = document.PresentationPart?.GetXDocument();
        XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        return xDoc?.Root?.Descendants(p14 + "sectionLst").Any() == true;
    }

    protected static bool SectionListUsesNumericSlideIds(FileInfo pptx)
    {
        using var document = PresentationDocument.Open(pptx.FullName, false);
        var xDoc = document.PresentationPart?.GetXDocument();
        XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var sections = xDoc?.Root?.Descendants(p14 + "section").ToList();
        if (sections is null || sections.Count == 0)
            return false;

        foreach (var section in sections)
        {
            if (section.Attribute("id") is null)
                return false;
            foreach (var slideRef in section.Descendants(p14 + "sldId"))
            {
                if (slideRef.Attribute(r + "id") is not null)
                    return false;
                if (!uint.TryParse((string?)slideRef.Attribute("id"), out _))
                    return false;
            }
        }
        return true;
    }

    protected static IReadOnlyList<string> GetSectionNames(FileInfo pptx)
    {
        using var document = PresentationDocument.Open(pptx.FullName, false);
        var xDoc = document.PresentationPart?.GetXDocument();
        XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        return xDoc?.Root?.Descendants(p14 + "section")
                .Select(section => (string?)section.Attribute("name") ?? string.Empty)
                .ToList()
            ?? [];
    }

    protected static void CorruptSectionListToUseRelationshipIds(FileInfo pptx)
    {
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var zip = ZipFile.Open(pptx.FullName, ZipArchiveMode.Update);
        var presentationEntry = zip.GetEntry("ppt/presentation.xml")!;
        XDocument xDoc;
        using (var stream = presentationEntry.Open())
            xDoc = XDocument.Load(stream);

        var slideIdToRelId = xDoc.Root!.Element(p + "sldIdLst")!
            .Elements(p + "sldId")
            .ToDictionary(e => (string)e.Attribute("id")!, e => (string)e.Attribute(r + "id")!, StringComparer.Ordinal);

        foreach (var section in xDoc.Root.Descendants(p14 + "section"))
            section.Attribute("id")?.Remove();

        foreach (var slideRef in xDoc.Root.Descendants(p14 + "sldId"))
        {
            var numericId = (string?)slideRef.Attribute("id");
            slideRef.Attribute("id")?.Remove();
            if (numericId is not null && slideIdToRelId.TryGetValue(numericId, out var relId))
                slideRef.SetAttributeValue(r + "id", relId);
        }

        presentationEntry.Delete();
        var newEntry = zip.CreateEntry("ppt/presentation.xml");
        using var writer = new StreamWriter(newEntry.Open());
        using var xmlWriter = System.Xml.XmlWriter.Create(writer);
        xDoc.WriteTo(xmlWriter);
    }

    protected static int CountSlides(FileInfo pptx)
    {
        using var document = PresentationDocument.Open(pptx.FullName, false);
        var presentationPart = document.PresentationPart;
        if (presentationPart is null)
            return 0;

        var presentation = presentationPart.Presentation;
        if (presentation is null)
            return 0;

        var slideIdList = presentation.SlideIdList;
        return slideIdList?.Count() ?? 0;
    }
}
