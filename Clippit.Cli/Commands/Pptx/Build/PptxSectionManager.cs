using System.Xml.Linq;
using Clippit;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Pptx.Build;

/// <summary>
/// Manages PPTX section XML (<p14:sectionLst>) for a presentation being built.
///
/// Usage pattern:
///   1. For each manifest section entry: call AddNewSection(name)
///   2. After AddSlidePart returns a dstSlidePart:
///        - if sections were loaded from source via LoadFrom: call RemapRelId(src, dst)
///        - otherwise: call AppendSlideToLastSection(dstRelId)
///   3. After builder.Dispose(): call SaveSectionsTo(destination)
/// </summary>
internal sealed class PptxSectionManager
{
    // Internally sections track destination relationship IDs while slides are copied.
    // SaveSectionsTo translates them to the numeric slide IDs used by <p14:sldId>.
    private readonly List<SectionInfo> _sections = [];

    // -------------------------------------------------------------------------
    // Building from manifest
    // -------------------------------------------------------------------------

    /// <summary>
    /// Adds a new named section. Slides subsequently appended via
    /// <see cref="AppendSlideToLastSection"/> belong to this section.
    /// </summary>
    public void AddNewSection(string? name)
    {
        _sections.Add(new SectionInfo(name ?? string.Empty, FormatGuid(Guid.NewGuid()), []));
    }

    /// <summary>
    /// Appends a copied slide's destination relationship ID to the last section.
    /// Call this after each <c>builder.AddSlidePart()</c> when NOT using LoadFrom.
    /// </summary>
    public void AppendSlideToLastSection(string dstRelId)
    {
        if (_sections.Count == 0)
            return;
        _sections[^1].SlideRelIds.Add(dstRelId);
    }

    // -------------------------------------------------------------------------
    // Loading from source document
    // -------------------------------------------------------------------------

    /// <summary>
    /// Attempts to load and append the section structure from a source presentation.
    /// Returns a session that remaps copied slide relationship IDs for this source.
    /// Returns null when the source has no sections.
    /// </summary>
    public SectionLoadSession? LoadFrom(PresentationDocument src)
    {
        var loaded = ReadSections(src);
        if (loaded.Count == 0)
            return null;

        var sectionInfos = loaded
            .Select(s => new SectionInfo(s.Name, FormatGuid(Guid.NewGuid()), s.SlideRelIds.ToList()))
            .ToList();
        _sections.AddRange(sectionInfos);
        return new SectionLoadSession(sectionInfos);
    }

    /// <summary>
    /// Reads the <c>&lt;p14:sectionLst&gt;</c> structure from a source presentation.
    /// Returns an empty list if no sections are defined.
    /// Each entry exposes source slide relationship IDs in section order so callers
    /// can map them while copying slides. The underlying XML uses numeric slide IDs.
    /// </summary>
    public static IReadOnlyList<(string Name, string Id, IReadOnlyList<string> SlideRelIds)> ReadSections(
        PresentationDocument src
    )
    {
        var presentationXDoc = src.PresentationPart?.GetXDocument();
        if (presentationXDoc is null)
            return [];

        // <p14:sectionLst> lives inside <p:ext> inside <p:extLst> on the root
        var sectionLst = presentationXDoc.Root?.Descendants(P14.sectionLst).FirstOrDefault();
        if (sectionLst is null)
            return [];

        var slideIdToRelId = presentationXDoc
            .Root?.Element(P.sldIdLst)
            ?.Elements(P.sldId)
            .Select(e => new
            {
                NumericId = (string?)e.Attribute(NoNamespace.id) ?? string.Empty,
                RelId = (string?)e.Attribute(R.id) ?? string.Empty,
            })
            .Where(x => x.NumericId.Length > 0 && x.RelId.Length > 0)
            .ToDictionary(x => x.NumericId, x => x.RelId, StringComparer.Ordinal);
        if (slideIdToRelId is null || slideIdToRelId.Count == 0)
            return [];

        var loaded = new List<(string Name, string Id, IReadOnlyList<string> SlideRelIds)>();
        foreach (var section in sectionLst.Elements(P14.section))
        {
            var name = (string?)section.Attribute(NoNamespace.name) ?? string.Empty;
            var id = (string?)section.Attribute(NoNamespace.id) ?? FormatGuid(Guid.NewGuid());
            var relIds = section
                .Descendants(P14.sldId)
                .Select(e => (string?)e.Attribute(NoNamespace.id) ?? string.Empty)
                .Select(id => slideIdToRelId.GetValueOrDefault(id) ?? string.Empty)
                .Where(id => id.Length > 0)
                .ToList();
            loaded.Add((name, id, relIds));
        }
        return loaded;
    }

    // -------------------------------------------------------------------------
    // Writing to destination
    // -------------------------------------------------------------------------

    /// <summary>
    /// Writes the accumulated <c>&lt;p14:sectionLst&gt;</c> into the destination
    /// presentation's extLst. Must be called after <c>builder.Dispose()</c> and
    /// before the destination <c>PresentationDocument</c> is disposed.
    /// Sections with no slides are omitted.
    /// </summary>
    public void SaveSectionsTo(PresentationDocument dst)
    {
        var nonEmpty = _sections.Where(s => s.SlideRelIds.Count > 0).ToList();
        if (nonEmpty.Count == 0)
            return;

        var presentationXDoc = dst.PresentationPart!.GetXDocument();
        var root = presentationXDoc.Root!;

        var relIdToSlideId = root.Element(P.sldIdLst)
            ?.Elements(P.sldId)
            .Select(e => new
            {
                RelId = (string?)e.Attribute(R.id) ?? string.Empty,
                NumericId = (string?)e.Attribute(NoNamespace.id) ?? string.Empty,
            })
            .Where(x => x.RelId.Length > 0 && x.NumericId.Length > 0)
            .ToDictionary(x => x.RelId, x => x.NumericId, StringComparer.Ordinal);
        if (relIdToSlideId is null || relIdToSlideId.Count == 0)
            return;

        // Build <p14:sectionLst>
        XNamespace p14ns = P14.p14;

        var sectionLst = new XElement(P14.sectionLst, new XAttribute(XNamespace.Xmlns + "p14", p14ns));

        foreach (var section in nonEmpty)
        {
            var sldIdLst = new XElement(P14.sldIdLst);
            foreach (var relId in section.SlideRelIds)
            {
                if (relIdToSlideId.TryGetValue(relId, out var numericId))
                    sldIdLst.Add(new XElement(P14.sldId, new XAttribute(NoNamespace.id, numericId)));
            }

            if (!sldIdLst.HasElements)
                continue;

            sectionLst.Add(
                new XElement(
                    P14.section,
                    new XAttribute(NoNamespace.name, section.Name),
                    new XAttribute(NoNamespace.id, section.Id),
                    sldIdLst
                )
            );
        }

        if (!sectionLst.HasElements)
            return;

        // Inject into extLst on the presentation root, creating it if absent
        var extLst = root.Element(P.extLst);
        if (extLst is null)
        {
            extLst = new XElement(P.extLst);
            root.Add(extLst);
        }

        // Remove any pre-existing sectionLst ext to avoid duplicates
        extLst.Elements(P.ext).Where(e => e.Descendants(P14.sectionLst).Any()).ToList().ForEach(e => e.Remove());

        extLst.Add(
            new XElement(P.ext, new XAttribute(NoNamespace.uri, "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"), sectionLst)
        );

        dst.PresentationPart!.PutXDocument(presentationXDoc);
    }

    private static string FormatGuid(Guid value) => value.ToString("B").ToUpperInvariant();

    internal sealed class SectionLoadSession(IReadOnlyList<SectionInfo> sections)
    {
        public void RemapRelId(string srcRelId, string dstRelId)
        {
            foreach (var section in sections)
            {
                var relIds = section.SlideRelIds;
                for (var i = 0; i < relIds.Count; i++)
                {
                    if (relIds[i] == srcRelId)
                        relIds[i] = dstRelId;
                }
            }
        }
    }

    internal sealed record SectionInfo(string Name, string Id, List<string> SlideRelIds);
}
