namespace Clippit.Internal;

internal static class Relationships
{
    internal static string GetNewRelationshipId()
    {
        var uid = Guid.NewGuid().ToString().Replace("-", "").AsSpan(0, 16);
        return string.Concat("rcId", uid);
    }
}
