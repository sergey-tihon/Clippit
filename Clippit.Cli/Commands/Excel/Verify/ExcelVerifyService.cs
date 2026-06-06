using Clippit.Cli.Commands.Common.Verify;
using Clippit.Cli.Infrastructure;
using Clippit.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Excel.Verify;

internal static class ExcelVerifyService
{
    public static VerifyResult Execute(InputSource input, FileFormatVersions officeVersion) =>
        VerifyExecutor.Execute(
            input,
            officeVersion,
            stream =>
            {
                using var document = SpreadsheetDocument.Open(
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
