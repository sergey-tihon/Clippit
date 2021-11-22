using System;

namespace Clippit.Internal;

internal static class Relationships
{
    internal static string GetNewRelationshipId() =>
        "rcId" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
}
