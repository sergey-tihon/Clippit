using System.Xml.Linq;

namespace Clippit.Word.Assembler;

internal static class FileDataExtensions
{
    internal static XElement GetBase64EncodedDocumentElement(this byte[] bytes)
    {
        var xmlString = $"<Document Data=\"{Convert.ToBase64String(bytes)}\" />";
        var sdt = new XElement(
            W.sdt,
            new XElement(W.sdtContent, new XElement(W.p, new XElement(W.r, new XElement(W.t, xmlString))))
        );

        return sdt;
    }
}
