using Clippit.Cli.Commands.Common.Verify;
using Clippit.Cli.Infrastructure;
using Clippit.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Word.Verify;

internal static class WordVerifyService
{
    public static VerifyResult Execute(InputSource input, FileFormatVersions officeVersion) =>
        VerifyExecutor.Execute(
            input,
            officeVersion,
            stream =>
            {
                using var document = WordprocessingDocument.Open(
                    stream,
                    isEditable: false,
                    new OpenSettings { AutoSave = false }
                );
                return OpenXmlPackageValidator.Validate(
                    document,
                    new OpenXmlValidationOptions { OfficeVersion = officeVersion }
                );
            }
        );
}
