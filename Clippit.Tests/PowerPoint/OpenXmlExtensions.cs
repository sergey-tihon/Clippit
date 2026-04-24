using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

public static class OpenXmlExtensions
{
    public static PresentationDocument OpenPresentation(Stream stream, bool isEditable, OpenSettings openSettings)
    {
        try
        {
            return PresentationDocument.Open(stream, isEditable, openSettings);
        }
        catch (OpenXmlPackageException e)
        {
            if (!e.ToString().Contains("Invalid Hyperlink"))
                throw;

            UriFixer.FixInvalidUri(stream, leaveOpen: true);
            stream.Position = 0;
            return PresentationDocument.Open(stream, isEditable, openSettings);
        }
    }
}
