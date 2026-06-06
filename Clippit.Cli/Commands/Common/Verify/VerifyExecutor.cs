using System.Xml;
using Clippit.Cli.Infrastructure;
using Clippit.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Common.Verify;

/// <summary>
/// Shared execution pipeline for the <c>pptx verify</c>, <c>word verify</c>, and
/// <c>excel verify</c> commands. Opens the input stream, runs a caller-supplied
/// validator delegate, converts package-open failures into a <c>package</c>
/// diagnostic on the result, and projects diagnostics into the CLI's
/// <see cref="VerifyDiagnostic"/> shape.
/// </summary>
internal static class VerifyExecutor
{
    public static VerifyResult Execute(
        InputSource input,
        FileFormatVersions officeVersion,
        Func<Stream, OpenXmlValidationResult> validate
    )
    {
        var inputStream = input.OpenSeekable();
        using (inputStream)
        {
            OpenXmlValidationResult validation;
            try
            {
                validation = validate(inputStream);
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
                validation = new OpenXmlValidationResult
                {
                    OfficeVersion = officeVersion,
                    Diagnostics =
                    [
                        new OpenXmlValidationDiagnostic
                        {
                            Kind = OpenXmlValidationDiagnosticKinds.Package,
                            Description = ex.Message,
                        },
                    ],
                };
            }

            return new VerifyResult
            {
                Input = input.DisplayName,
                OfficeVersion = validation.OfficeVersion.ToString(),
                Valid = validation.Valid,
                Diagnostics = validation.Diagnostics.Select(ToVerifyDiagnostic).ToList(),
            };
        }
    }

    private static VerifyDiagnostic ToVerifyDiagnostic(OpenXmlValidationDiagnostic diagnostic) =>
        new()
        {
            Kind = diagnostic.Kind,
            Code = diagnostic.Code,
            Description = diagnostic.Description,
            Part = diagnostic.Part,
            Path = diagnostic.Path,
            Element = diagnostic.Element,
            Attribute = diagnostic.Attribute,
            RelationshipId = diagnostic.RelationshipId,
        };
}
