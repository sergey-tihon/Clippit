using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Excel;

public class SmlDocument : OpenXmlPowerToolsDocument
{
    private const string NotSpreadsheetExceptionMessage = "The document is not a SpreadsheetML document.";

    public SmlDocument(OpenXmlPowerToolsDocument original)
        : base(original)
    {
        if (GetDocumentType() != typeof(SpreadsheetDocument))
            throw new PowerToolsDocumentException(NotSpreadsheetExceptionMessage);
    }

    public SmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
        : base(original, convertToTransitional)
    {
        if (GetDocumentType() != typeof(SpreadsheetDocument))
            throw new PowerToolsDocumentException(NotSpreadsheetExceptionMessage);
    }

    public SmlDocument(string fileName)
        : base(fileName)
    {
        if (GetDocumentType() != typeof(SpreadsheetDocument))
            throw new PowerToolsDocumentException(NotSpreadsheetExceptionMessage);
    }

    public SmlDocument(string fileName, bool convertToTransitional)
        : base(fileName, convertToTransitional)
    {
        if (GetDocumentType() != typeof(SpreadsheetDocument))
            throw new PowerToolsDocumentException(NotSpreadsheetExceptionMessage);
    }

    public SmlDocument(string fileName, byte[] byteArray)
        : base(byteArray)
    {
        FileName = fileName;
        if (GetDocumentType() != typeof(SpreadsheetDocument))
            throw new PowerToolsDocumentException(NotSpreadsheetExceptionMessage);
    }

    public SmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
        : base(byteArray, convertToTransitional)
    {
        FileName = fileName;
        if (GetDocumentType() != typeof(SpreadsheetDocument))
            throw new PowerToolsDocumentException(NotSpreadsheetExceptionMessage);
    }

    public SmlDocument(string fileName, MemoryStream memStream)
        : base(fileName, memStream) { }

    public SmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
        : base(fileName, memStream, convertToTransitional) { }

    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public XElement ConvertToHtml(SmlToHtmlConverterSettings htmlConverterSettings, string tableName) =>
        SmlToHtmlConverter.ConvertTableToHtml(this, htmlConverterSettings, tableName);

    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public XElement ConvertTableToHtml(string tableName) =>
        SmlToHtmlConverter.ConvertTableToHtml(this, new SmlToHtmlConverterSettings(), tableName);
}
