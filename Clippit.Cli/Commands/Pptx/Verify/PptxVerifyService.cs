using Clippit.Cli.Commands.Common.Verify;
using Clippit.Cli.Infrastructure;
using Clippit.Core;
using Clippit.PowerPoint;
using DocumentFormat.OpenXml;

namespace Clippit.Cli.Commands.Pptx.Verify;

internal static class PptxVerifyService
{
    public static VerifyResult Execute(InputSource input, FileFormatVersions officeVersion) =>
        VerifyExecutor.Execute(
            input,
            officeVersion,
            stream =>
                PresentationValidator.Validate(stream, new OpenXmlValidationOptions { OfficeVersion = officeVersion })
        );
}
