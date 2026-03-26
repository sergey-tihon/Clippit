using System.Security.Cryptography;
using System.Text;

namespace Clippit.Internal;

internal static class Relationships
{
    /// <summary>
    /// Generates a deterministic relationship ID from a context string.
    /// The same context always produces the same ID, ensuring stable output across runs.
    /// The context must be unique enough to avoid collisions (e.g., combine source part URI and old relationship ID).
    /// </summary>
    internal static string GetNewRelationshipId(string context)
    {
        var hash = SHA256.HashData(Encoding.UTF8.GetBytes(context));
        return "rId" + Convert.ToHexString(hash)[..20];
    }

    /// <summary>
    /// Generates a deterministic relationship ID from raw byte content (e.g., image bytes).
    /// </summary>
    internal static string GetNewRelationshipId(ReadOnlySpan<byte> content)
    {
        var hash = SHA256.HashData(content);
        return "rId" + Convert.ToHexString(hash)[..20];
    }
}
