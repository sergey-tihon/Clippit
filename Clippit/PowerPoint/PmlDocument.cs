using Clippit.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint;

public partial class PmlDocument : OpenXmlPowerToolsDocument
{
    private const string NotPresentationExceptionMessage = "The document is not a PresentationML document.";

    public PmlDocument(OpenXmlPowerToolsDocument original)
        : base(original)
    {
        if (GetDocumentType() != typeof(PresentationDocument))
            throw new PowerToolsDocumentException(NotPresentationExceptionMessage);
    }

    public PmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
        : base(original, convertToTransitional)
    {
        if (GetDocumentType() != typeof(PresentationDocument))
            throw new PowerToolsDocumentException(NotPresentationExceptionMessage);
    }

    public PmlDocument(string fileName)
        : base(fileName)
    {
        if (GetDocumentType() != typeof(PresentationDocument))
            throw new PowerToolsDocumentException(NotPresentationExceptionMessage);
    }

    public PmlDocument(string fileName, bool convertToTransitional)
        : base(fileName, convertToTransitional)
    {
        if (GetDocumentType() != typeof(PresentationDocument))
            throw new PowerToolsDocumentException(NotPresentationExceptionMessage);
    }

    public PmlDocument(string fileName, byte[] byteArray)
        : base(byteArray)
    {
        FileName = fileName;
        if (GetDocumentType() != typeof(PresentationDocument))
            throw new PowerToolsDocumentException(NotPresentationExceptionMessage);
    }

    public PmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
        : base(byteArray, convertToTransitional)
    {
        FileName = fileName;
        if (GetDocumentType() != typeof(PresentationDocument))
            throw new PowerToolsDocumentException(NotPresentationExceptionMessage);
    }

    public PmlDocument(string fileName, MemoryStream memStream)
        : base(fileName, memStream) { }

    public PmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
        : base(fileName, memStream, convertToTransitional) { }

    public PmlDocument SearchAndReplace(string search, string replace, bool matchCase)
    {
        return TextReplacer.SearchAndReplace(this, search, replace, matchCase);
    }
}
