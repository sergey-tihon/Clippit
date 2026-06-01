namespace Clippit.Core;

/// <summary>
/// Describes a validation diagnostic found in an OpenXml package.
/// </summary>
public sealed class OpenXmlValidationDiagnostic
{
    /// <summary>
    /// Stable diagnostic category, such as <c>schema</c>, <c>relationship</c>, or a document-specific kind.
    /// </summary>
    public required string Kind { get; init; }

    /// <summary>
    /// Stable validator-specific diagnostic code, when available.
    /// </summary>
    public string? Code { get; init; }

    /// <summary>
    /// Human-readable validation message.
    /// </summary>
    public required string Description { get; init; }

    /// <summary>
    /// OpenXml part URI related to the diagnostic, when available.
    /// </summary>
    public string? Part { get; init; }

    /// <summary>
    /// OpenXml validator XPath related to the diagnostic, when available.
    /// </summary>
    public string? Path { get; init; }

    /// <summary>
    /// Element name related to the diagnostic, when available. Schema diagnostics use
    /// <c>{namespace}localName</c>; relationship diagnostics use the local name.
    /// </summary>
    public string? Element { get; init; }

    /// <summary>
    /// Attribute local name related to the diagnostic, when available.
    /// </summary>
    public string? Attribute { get; init; }

    /// <summary>
    /// Relationship ID related to the diagnostic, when available.
    /// </summary>
    public string? RelationshipId { get; init; }
}
