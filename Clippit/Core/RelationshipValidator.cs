// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Core;

/// <summary>
/// Describes a single relationship validation error found in an OpenXml package.
/// </summary>
/// <param name="PartUri">URI of the part that owns the problematic relationship attribute.</param>
/// <param name="ElementName">XML element name that carries the dangling relationship ID.</param>
/// <param name="AttributeName">XML attribute name that holds the relationship ID.</param>
/// <param name="RelationshipId">The relationship ID value that could not be resolved.</param>
/// <param name="Description">Human-readable description of the problem.</param>
public sealed record RelationshipValidationError(
    Uri PartUri,
    XName ElementName,
    XName AttributeName,
    string RelationshipId,
    string Description
);

/// <summary>
/// Validates that every relationship ID referenced inside XML markup can be resolved
/// to an actual relationship registered on that part.
/// </summary>
/// <remarks>
/// <para>
/// The built-in <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> checks
/// element-level schema conformance but does not detect <em>dangling</em> relationship
/// references — attributes such as <c>r:id</c>, <c>r:embed</c>, or <c>r:link</c> whose
/// value does not correspond to any relationship registered with the part.
/// </para>
/// <para>
/// A dangling reference causes a <see cref="KeyNotFoundException"/> at runtime when code
/// tries to resolve it (e.g. during slide copying or publishing). Running this validator
/// before such operations lets callers detect the problem early and skip or repair the
/// affected element instead of crashing.
/// </para>
/// </remarks>
public static class RelationshipValidator
{
    /// <summary>
    /// XML attribute names that carry relationship IDs in OpenXml markup
    /// (namespace <c>http://schemas.openxmlformats.org/officeDocument/2006/relationships</c>).
    /// </summary>
    private static readonly XName[] s_relationshipAttributeNames =
    [
        R.embed,
        R.link,
        R.id,
        R.cs,
        R.dm,
        R.lo,
        R.qs,
        R.href,
        R.pict,
        R.blip,
    ];

    /// <summary>
    /// Validates that every relationship attribute value in every XML part of the
    /// given package resolves to a registered relationship on that part.
    /// </summary>
    /// <param name="package">The package to validate.</param>
    /// <returns>
    /// A sequence of <see cref="RelationshipValidationError"/> items describing each
    /// unresolvable relationship reference; empty when the package is clean.
    /// </returns>
    public static IEnumerable<RelationshipValidationError> Validate(OpenXmlPackage package)
    {
        ArgumentNullException.ThrowIfNull(package);

        var errors = new List<RelationshipValidationError>();

        foreach (var part in package.GetAllParts())
        {
            // Only XML parts carry markup with relationship attributes.
            if (!part.ContentType.EndsWith("xml", StringComparison.OrdinalIgnoreCase))
                continue;

            XDocument xDoc;
            try
            {
                xDoc = part.GetXDocument();
            }
            catch
            {
                // Skip parts we cannot parse.
                continue;
            }

            // Build a set of all relationship IDs registered on this part.
            var registeredIds = BuildRegisteredRelationshipIds(part);

            foreach (var element in xDoc.Descendants())
            {
                foreach (var attrName in s_relationshipAttributeNames)
                {
                    var attr = element.Attribute(attrName);
                    if (attr is null || string.IsNullOrEmpty(attr.Value))
                        continue;

                    if (!registeredIds.Contains(attr.Value))
                    {
                        errors.Add(
                            new RelationshipValidationError(
                                part.Uri,
                                element.Name,
                                attrName,
                                attr.Value,
                                $"Part '{part.Uri}': element '{element.Name.LocalName}' attribute '{attrName.LocalName}' references relationship ID '{attr.Value}' which is not registered on this part."
                            )
                        );
                    }
                }
            }
        }

        return errors;
    }

    /// <summary>
    /// Returns <see langword="true"/> when no dangling relationship references are found.
    /// </summary>
    public static bool IsValid(OpenXmlPackage package) => !Validate(package).Any();

    private static HashSet<string> BuildRegisteredRelationshipIds(OpenXmlPart part)
    {
        var ids = new HashSet<string>(StringComparer.Ordinal);

        foreach (var pair in part.Parts)
            ids.Add(pair.RelationshipId);

        foreach (var rel in part.ExternalRelationships)
            ids.Add(rel.Id);

        foreach (var rel in part.HyperlinkRelationships)
            ids.Add(rel.Id);

        foreach (var rel in part.DataPartReferenceRelationships)
            ids.Add(rel.Id);

        return ids;
    }
}
