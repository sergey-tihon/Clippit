using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Clippit.Word.Assembler
{
    internal static class FileDataExtensions
    {
        internal static XElement GetBase64EncodedDocumentElement(this byte[] bytes)
        {
            // return a Content Control with byte[] base64 encoded
            /*<w:sdt><w:sdtContent>
                <w:p w:rsidR="00FB5781" w:rsidRDefault="00AE77E8">
                <w:r><w:t>&lt;Document Data="BASE64" /&gt;</w:t></w:r></w:p></w:sdtContent></w:sdt>
            */

            var xmlString = $"<Document Data=\"{Convert.ToBase64String(bytes)}\" />";
            XElement sdt = new XElement(
                W.sdt,
                new XElement(W.sdtContent, new XElement(W.p, new XElement(W.r, new XElement(W.t, xmlString))))
            );

            return sdt;
        }
    }
}
