using System.Xml.Linq;

namespace Clippit.Word.Assembler
{
    internal class RunReplacementInfo
    {
        public XElement Xml { get; set; }
        public string XmlExceptionMessage { get; set; }
        public string SchemaValidationMessage { get; set; }
    }
}
