using Clippit.Cli.Infrastructure;
using Clippit.Core;
using Clippit.PowerPoint;
using DocumentFormat.OpenXml;

namespace Clippit.Cli.Commands.Pptx.Verify;

internal static class PptxVerifyService
{
    public static VerifyResult Execute(InputSource input, FileFormatVersions officeVersion)
    {
        var inputStream = input.OpenSeekable();
        using (inputStream)
        {
            var validation = PresentationValidator.Validate(
                inputStream,
                new OpenXmlValidationOptions { OfficeVersion = officeVersion }
            );
            var diagnostics = validation.Diagnostics.Select(ToVerifyDiagnostic).ToList();

            return new VerifyResult
            {
                Input = input.DisplayName,
                OfficeVersion = validation.OfficeVersion.ToString(),
                Valid = validation.Valid,
                Diagnostics = diagnostics,
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
