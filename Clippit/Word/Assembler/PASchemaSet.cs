using System.Runtime.CompilerServices;
using System.Xml;
using System.Xml.Schema;

namespace Clippit.Word.Assembler
{
    internal class PASchemaSet
    {
        private readonly string XsdMarkup;
        public readonly XmlSchemaSet SchemaSet;

        internal PASchemaSet(string xsdMarkup)
        {
            this.XsdMarkup = xsdMarkup;
            this.SchemaSet = new XmlSchemaSet();

            XmlSchema schema = XmlSchema.Read(XmlReader.Create(new StringReader(XsdMarkup)), null);
            this.SchemaSet.Add(schema);
        }
    }
}
