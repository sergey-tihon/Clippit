using System;

namespace Clippit.Internal;

internal static class Relationships
{
    internal static string GetNewRelationshipId() =>
        string.Concat("rcId", Guid.NewGuid().ToString().Replace("-", "").AsSpan(0, 16));
}
