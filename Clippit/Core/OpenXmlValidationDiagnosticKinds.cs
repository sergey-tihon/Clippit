namespace Clippit.Core;

/// <summary>
/// Common validation diagnostic kinds used by OpenXml package validators.
/// </summary>
public static class OpenXmlValidationDiagnosticKinds
{
    public const string Package = "package";
    public const string Schema = "schema";
    public const string Semantic = "semantic";
    public const string MarkupCompatibility = "markupCompatibility";
    public const string Relationship = "relationship";
    public const string Unknown = "unknown";
}
