using Clippit.Core;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint;

/// <summary>
/// Validates PowerPoint presentations using generic OpenXml checks plus PowerPoint-specific package rules.
/// </summary>
public static class PresentationValidator
{
    /// <summary>
    /// Opens and validates a PowerPoint presentation stream.
    /// </summary>
    /// <remarks>
    /// The caller owns <paramref name="stream" /> and is responsible for disposing it.
    /// </remarks>
    public static OpenXmlValidationResult Validate(Stream stream, OpenXmlValidationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(stream);

        options ??= new OpenXmlValidationOptions();

        try
        {
            using var document = PresentationDocument.Open(stream, false, new OpenSettings { AutoSave = false });
            return Validate(document, options);
        }
        catch (Exception ex)
            when (ex
                    is OpenXmlPackageException
                        or FileFormatException
                        or InvalidDataException
                        or System.Xml.XmlException
                        or FormatException
            )
        {
            return InvalidPackage(options, ex.Message);
        }
    }

    /// <summary>
    /// Validates an already-open PowerPoint presentation.
    /// </summary>
    public static OpenXmlValidationResult Validate(
        PresentationDocument document,
        OpenXmlValidationOptions? options = null
    )
    {
        ArgumentNullException.ThrowIfNull(document);

        options ??= new OpenXmlValidationOptions();

        var diagnostics = new List<OpenXmlValidationDiagnostic>();

        if (document.PresentationPart is null)
        {
            diagnostics.Add(
                new OpenXmlValidationDiagnostic
                {
                    Kind = OpenXmlValidationDiagnosticKinds.Package,
                    Description = "Package does not contain a presentation part.",
                }
            );

            return new OpenXmlValidationResult { OfficeVersion = options.OfficeVersion, Diagnostics = diagnostics };
        }

        diagnostics.AddRange(OpenXmlPackageValidator.Validate(document, options).Diagnostics);
        diagnostics.AddRange(PresentationSectionValidator.Validate(document, options));

        return new OpenXmlValidationResult { OfficeVersion = options.OfficeVersion, Diagnostics = diagnostics };
    }

    private static OpenXmlValidationResult InvalidPackage(OpenXmlValidationOptions options, string description) =>
        new()
        {
            OfficeVersion = options.OfficeVersion,
            Diagnostics =
            [
                new OpenXmlValidationDiagnostic
                {
                    Kind = OpenXmlValidationDiagnosticKinds.Package,
                    Description = description,
                },
            ],
        };
}
