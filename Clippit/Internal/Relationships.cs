using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Internal;

internal static class Relationships
{
    /// <summary>
    /// Generates a random relationship ID using GUID.
    /// This is the original non-deterministic approach used by Word, Excel, and HTML converters.
    /// </summary>
    internal static string GetNewRelationshipId()
    {
        var uid = Guid.NewGuid().ToString().Replace("-", "").AsSpan(0, 16);
        return string.Concat("rcId", uid);
    }
}

/// <summary>
/// Generates relationship IDs in the format Office uses: rId1, rId2, rId3, …
/// Initialise once per part by scanning its existing relationships to find the
/// highest rIdN already in use, then increment for every new ID needed — O(1) per call.
/// </summary>
internal sealed class RelationshipIdGenerator
{
    private int _next;

    /// <summary>
    /// Initialise from an <see cref="OpenXmlPart"/>, scanning all of its
    /// existing relationship collections once to find the current high-water mark.
    /// </summary>
    internal RelationshipIdGenerator(OpenXmlPart part)
    {
        _next =
            ComputeMax(
                part.Parts.Select(p => p.RelationshipId)
                    .Concat(part.ExternalRelationships.Select(r => r.Id))
                    .Concat(part.HyperlinkRelationships.Select(r => r.Id))
            ) + 1;
    }

    /// <summary>
    /// Initialise from an explicit set of already-used IDs (e.g. for raw ZIP/.rels processing).
    /// </summary>
    internal RelationshipIdGenerator(IEnumerable<string> existingIds)
    {
        _next = ComputeMax(existingIds) + 1;
    }

    /// <summary>
    /// Start from rId1 with no existing IDs to scan (e.g. for a brand-new part).
    /// </summary>
    internal RelationshipIdGenerator()
    {
        _next = 1;
    }

    /// <summary>Returns the next available relationship ID and advances the counter.</summary>
    internal string Next() => $"rId{_next++}";

    private static int ComputeMax(IEnumerable<string> ids)
    {
        var max = 0;
        foreach (var id in ids)
        {
            if (
                id.Length > 3
                && id.StartsWith("rId", StringComparison.Ordinal)
                && int.TryParse(id.AsSpan(3), out var n)
                && n > max
            )
                max = n;
        }
        return max;
    }
}

internal static class RelationshipIdGeneratorExtensions
{
    /// <summary>
    /// Creates a new <see cref="RelationshipIdGenerator"/> for <paramref name="part"/> by
    /// scanning its current relationships once to find the highest rIdN in use.
    /// Use this when you need a single fresh ID and SDK methods may have added relationships
    /// since the last generator was created.
    /// For hot loops that add many IDs to the same part without SDK interleaving, call
    /// <see cref="GetCachedRelationshipIdGenerator"/> instead.
    /// </summary>
    internal static RelationshipIdGenerator GetRelationshipIdGenerator(this OpenXmlPart part) =>
        new(part);

    /// <summary>
    /// Returns a cached <see cref="RelationshipIdGenerator"/> for <paramref name="part"/>,
    /// creating it on first call (one scan) and reusing it on subsequent calls — O(1).
    /// Only safe to use when no other code adds relationships to the part between calls
    /// (i.e. all additions go through this generator).
    /// </summary>
    internal static RelationshipIdGenerator GetCachedRelationshipIdGenerator(this OpenXmlPart part)
    {
        var gen = part.Annotation<RelationshipIdGenerator>();
        if (gen is null)
        {
            gen = new RelationshipIdGenerator(part);
            part.AddAnnotation(gen);
        }
        return gen;
    }
}
