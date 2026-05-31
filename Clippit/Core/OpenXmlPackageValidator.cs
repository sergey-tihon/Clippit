using System.Xml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace Clippit.Core;

/// <summary>
/// Validates generic OpenXml package rules that apply across Word, Excel, and PowerPoint files.
/// </summary>
public static class OpenXmlPackageValidator
{
    /// <summary>
    /// Validates schema conformance and dangling relationship references for an OpenXml package.
    /// </summary>
    public static OpenXmlValidationResult Validate(OpenXmlPackage package, OpenXmlValidationOptions? options = null)
    {
        ArgumentNullException.ThrowIfNull(package);

        options ??= new OpenXmlValidationOptions();

        var diagnostics = new List<OpenXmlValidationDiagnostic>();
        var validator = new OpenXmlValidator(options.OfficeVersion);

        try
        {
            diagnostics.AddRange(validator.Validate(package).Select(ToDiagnostic));
        }
        catch (Exception ex)
            when (ex
                    is OpenXmlPackageException
                        or FileFormatException
                        or InvalidDataException
                        or XmlException
                        or FormatException
            )
        {
            var diagnostic = new OpenXmlValidationDiagnostic
            {
                Kind = OpenXmlValidationDiagnosticKinds.Package,
                Description = ex.Message,
            };
            diagnostics.Add(diagnostic);
        }
        diagnostics.AddRange(RelationshipValidator.Validate(package).Select(ToDiagnostic));

        return new OpenXmlValidationResult { OfficeVersion = options.OfficeVersion, Diagnostics = diagnostics };
    }

    private static OpenXmlValidationDiagnostic ToDiagnostic(ValidationErrorInfo error) =>
        new()
        {
            Kind = ToKind(error.ErrorType),
            Code = error.Id,
            Description = error.Description,
            Part = error.Part?.Uri.ToString(),
            Path = error.Path?.XPath,
            Element = error.Node is null ? null : $"{{{error.Node.NamespaceUri}}}{error.Node.LocalName}",
        };

    private static OpenXmlValidationDiagnostic ToDiagnostic(RelationshipValidationError error) =>
        new()
        {
            Kind = error.Kind,
            Description = error.Description,
            Part = error.PartUri.ToString(),
            Element = error.ElementName.LocalName,
            Attribute = error.AttributeName.LocalName,
            RelationshipId = error.RelationshipId,
        };

    private static string ToKind(ValidationErrorType errorType) =>
        errorType switch
        {
            ValidationErrorType.Schema => OpenXmlValidationDiagnosticKinds.Schema,
            ValidationErrorType.Semantic => OpenXmlValidationDiagnosticKinds.Semantic,
            ValidationErrorType.Package => OpenXmlValidationDiagnosticKinds.Package,
            ValidationErrorType.MarkupCompatibility => OpenXmlValidationDiagnosticKinds.MarkupCompatibility,
            _ => OpenXmlValidationDiagnosticKinds.Unknown,
        };
}
