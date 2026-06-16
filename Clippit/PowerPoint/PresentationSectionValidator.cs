using System.Xml;
using System.Xml.Linq;
using Clippit.Core;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint;

internal static class PresentationSectionValidator
{
    public static IEnumerable<OpenXmlValidationDiagnostic> Validate(
        PresentationDocument document,
        OpenXmlValidationOptions? options = null
    )
    {
        var presentationPart = document.PresentationPart;
        if (presentationPart is null)
            yield break;

        // GetXDocument() applies StrictTranslatingXmlReader so that Strict namespace URIs
        // (purl.oclc.org) are transparently mapped to Transitional ones before the XDocument
        // is built.  Without this, P.sldIdLst lookups return nothing on Strict files and
        // every section slide reference would appear as a false-positive validation error.
        XDocument? presentationXDoc = null;
        OpenXmlValidationDiagnostic? parseDiagnostic = null;
        try
        {
            presentationXDoc = presentationPart.GetXDocument();
        }
        catch (XmlException ex)
        {
            parseDiagnostic = new OpenXmlValidationDiagnostic
            {
                Kind = OpenXmlValidationDiagnosticKinds.Package,
                Description = $"Part '{presentationPart.Uri}' contains malformed XML: {ex.Message}",
                Part = presentationPart.Uri.ToString(),
                Element = "presentation",
            };
        }

        if (parseDiagnostic is not null)
        {
            yield return parseDiagnostic;
            yield break;
        }

        var root = presentationXDoc?.Root;
        if (root is null)
            yield break;

        var slideIds = root.Element(P.sldIdLst)
            ?.Elements(P.sldId)
            .Select(e => (string?)e.Attribute(NoNamespace.id) ?? string.Empty)
            .Where(id => id.Length > 0)
            .ToHashSet(StringComparer.Ordinal);
        slideIds ??= [];

        foreach (var section in root.Descendants(P14.section))
        {
            var sectionName = (string?)section.Attribute(NoNamespace.name) ?? string.Empty;
            var sectionId = (string?)section.Attribute(NoNamespace.id);
            if (string.IsNullOrWhiteSpace(sectionId))
            {
                yield return Error(
                    $"PPTX section '{sectionName}' is missing required section id GUID.",
                    "section",
                    attribute: "id"
                );
            }
            else if (!Guid.TryParseExact(sectionId, "B", out _))
            {
                yield return Error(
                    $"PPTX section '{sectionName}' has invalid section id '{sectionId}'.",
                    "section",
                    attribute: "id"
                );
            }

            foreach (var sldId in section.Descendants(P14.sldId))
            {
                var numericId = (string?)sldId.Attribute(NoNamespace.id);
                var relId = (string?)sldId.Attribute(R.id);

                if (!string.IsNullOrEmpty(relId))
                {
                    yield return Error(
                        $"PPTX section '{sectionName}' references slide by relationship id '{relId}'. Sections must use numeric slide ids from p:sldIdLst.",
                        "sldId",
                        relationshipId: relId
                    );
                }

                if (string.IsNullOrWhiteSpace(numericId))
                {
                    yield return Error(
                        $"PPTX section '{sectionName}' contains a slide reference without numeric id.",
                        "sldId",
                        attribute: "id",
                        relationshipId: relId
                    );
                    continue;
                }

                if (!uint.TryParse(numericId, out _))
                {
                    yield return Error(
                        $"PPTX section '{sectionName}' contains non-numeric slide id '{numericId}'.",
                        "sldId",
                        attribute: "id"
                    );
                    continue;
                }

                if (!slideIds.Contains(numericId))
                {
                    yield return Error(
                        $"PPTX section '{sectionName}' references slide id '{numericId}' that is not present in p:sldIdLst.",
                        "sldId",
                        attribute: "id"
                    );
                }
            }
        }
    }

    private static OpenXmlValidationDiagnostic Error(
        string description,
        string element,
        string? attribute = null,
        string? relationshipId = null
    ) =>
        new()
        {
            Kind = PresentationValidationDiagnosticKinds.Section,
            Description = description,
            Part = "/ppt/presentation.xml",
            Element = element,
            Attribute = attribute,
            RelationshipId = relationshipId,
        };
}
