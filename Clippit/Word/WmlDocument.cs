using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using Clippit.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Word;

public partial class WmlDocument : OpenXmlPowerToolsDocument
{
    private const string NotWordprocessingExceptionMessage =
        "The document is not a WordprocessingML document.";

    public WmlDocument(OpenXmlPowerToolsDocument original)
        : base(original)
    {
        if (GetDocumentType() != typeof(WordprocessingDocument))
            throw new PowerToolsDocumentException(NotWordprocessingExceptionMessage);
    }

    public WmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
        : base(original, convertToTransitional)
    {
        if (GetDocumentType() != typeof(WordprocessingDocument))
            throw new PowerToolsDocumentException(NotWordprocessingExceptionMessage);
    }

    public WmlDocument(string fileName)
        : base(fileName)
    {
        if (GetDocumentType() != typeof(WordprocessingDocument))
            throw new PowerToolsDocumentException(NotWordprocessingExceptionMessage);
    }

    public WmlDocument(string fileName, bool convertToTransitional)
        : base(fileName, convertToTransitional)
    {
        if (GetDocumentType() != typeof(WordprocessingDocument))
            throw new PowerToolsDocumentException(NotWordprocessingExceptionMessage);
    }

    public WmlDocument(string fileName, byte[] byteArray)
        : base(byteArray)
    {
        FileName = fileName;
        if (GetDocumentType() != typeof(WordprocessingDocument))
            throw new PowerToolsDocumentException(NotWordprocessingExceptionMessage);
    }

    public WmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
        : base(byteArray, convertToTransitional)
    {
        FileName = fileName;
        if (GetDocumentType() != typeof(WordprocessingDocument))
            throw new PowerToolsDocumentException(NotWordprocessingExceptionMessage);
    }

    public WmlDocument(string fileName, MemoryStream memStream)
        : base(fileName, memStream) { }

    public WmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
        : base(fileName, memStream, convertToTransitional) { }

    public PtMainDocumentPart MainDocumentPart
    {
        get
        {
            using var ms = new MemoryStream(this.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var partElement = wDoc.MainDocumentPart.GetXDocument().Root;
            var childNodes = partElement.Nodes().ToList();
            foreach (var item in childNodes)
                item.Remove();
            return new PtMainDocumentPart(
                this,
                wDoc.MainDocumentPart.Uri,
                partElement.Name,
                partElement.Attributes(),
                childNodes
            );
        }
    }

    public WmlDocument(WmlDocument other, params XElement[] replacementParts)
        : base(other)
    {
        using var streamDoc = new OpenXmlMemoryStreamDocument(this);
        using (var package = streamDoc.GetPackage())
        {
            foreach (var replacementPart in replacementParts)
            {
                var uriAttribute = replacementPart.Attribute(PtOpenXml.Uri);
                if (uriAttribute == null)
                    throw new OpenXmlPowerToolsException(
                        "Replacement part does not contain a Uri as an attribute"
                    );
                var uri = uriAttribute.Value;
                var part = package.GetParts().FirstOrDefault(p => p.Uri.ToString() == uri);
                using var partStream = part.GetStream(FileMode.Create, FileAccess.Write);
                using var partXmlWriter = XmlWriter.Create(partStream);
                replacementPart.Save(partXmlWriter);
            }
        }
        this.DocumentByteArray = streamDoc.GetModifiedDocument().DocumentByteArray;
    }

    public WmlDocument SearchAndReplace(string search, string replace, bool matchCase)
    {
        return TextReplacer.SearchAndReplace(this, search, replace, matchCase);
    }
}

public class PtMainDocumentPart : XElement
{
    private readonly WmlDocument ParentWmlDocument;

    public PtWordprocessingCommentsPart WordprocessingCommentsPart
    {
        get
        {
            using var ms = new MemoryStream(ParentWmlDocument.DocumentByteArray);
            using var wDoc = WordprocessingDocument.Open(ms, false);
            var commentsPart = wDoc.MainDocumentPart.WordprocessingCommentsPart;
            if (commentsPart == null)
                return null;
            var partElement = commentsPart.GetXDocument().Root;
            var childNodes = partElement.Nodes().ToList();
            foreach (var item in childNodes)
                item.Remove();
            return new PtWordprocessingCommentsPart(
                this.ParentWmlDocument,
                commentsPart.Uri,
                partElement.Name,
                partElement.Attributes(),
                childNodes
            );
        }
    }

    public PtMainDocumentPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
        : base(name, values)
    {
        ParentWmlDocument = wmlDocument;
        this.Add(
            new XAttribute(PtOpenXml.Uri, uri),
            new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt)
        );
    }
}

public class PtWordprocessingCommentsPart : XElement
{
    private WmlDocument ParentWmlDocument;

    public PtWordprocessingCommentsPart(
        WmlDocument wmlDocument,
        Uri uri,
        XName name,
        params object[] values
    )
        : base(name, values)
    {
        ParentWmlDocument = wmlDocument;
        this.Add(
            new XAttribute(PtOpenXml.Uri, uri),
            new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt)
        );
    }
}
