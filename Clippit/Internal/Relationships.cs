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
