using System.IO;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint;

public static partial class PresentationBuilder
{
    public static IFluentPresentationBuilder Create(PresentationDocument document)
    {
        return new FluentPresentationBuilder(document);
    }

    public static PresentationDocument NewDocument(Stream stream)
    {
        return PresentationDocument.Create(stream, DocumentFormat.OpenXml.PresentationDocumentType.Presentation, false);
    }
}
