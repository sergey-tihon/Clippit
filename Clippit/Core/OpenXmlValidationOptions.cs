using DocumentFormat.OpenXml;

namespace Clippit.Core;

/// <summary>
/// Options for OpenXml package validation.
/// </summary>
public sealed class OpenXmlValidationOptions
{
    /// <summary>
    /// OpenXml schema version to validate against.
    /// </summary>
    public FileFormatVersions OfficeVersion { get; init; } = FileFormatVersions.Microsoft365;
}
