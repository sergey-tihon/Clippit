using Clippit;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

public static class OpenXmlExtensions
{
    public static PresentationDocument OpenPresentation(Stream stream, bool isEditable, OpenSettings openSettings) =>
        UriFixer.OpenPresentationDocument(stream, isEditable, openSettings);
}
