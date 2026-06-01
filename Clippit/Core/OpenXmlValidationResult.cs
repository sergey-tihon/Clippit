using DocumentFormat.OpenXml;

namespace Clippit.Core;

/// <summary>
/// Result of OpenXml package validation.
/// </summary>
public sealed class OpenXmlValidationResult
{
    /// <summary>
    /// OpenXml schema version used for validation.
    /// </summary>
    public required FileFormatVersions OfficeVersion { get; init; }

    /// <summary>
    /// Validation diagnostics found in the package.
    /// </summary>
    public required IReadOnlyList<OpenXmlValidationDiagnostic> Diagnostics { get; init; }

    /// <summary>
    /// Returns <see langword="true" /> when no diagnostics were found.
    /// </summary>
    public bool Valid => Diagnostics.Count == 0;
}
