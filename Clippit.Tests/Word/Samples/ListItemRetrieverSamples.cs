using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word.Samples
{
    public class ListItemRetrieverSamples() : Clippit.Tests.TestsBase
    {
        private class XmlStackItem
        {
            public XElement Element { get; init; }
            public int[] LevelNumbers { get; init; }
        }

        [Test]
        public void Sample()
        {
            using var wDoc = WordprocessingDocument.Open(
                "../../../Word/Samples/ListItemRetriever/NumberedListTest.docx",
                false
            );
            var abstractNumId = 0;
            var xml = ConvertDocToXml(wDoc, abstractNumId);
            Console.WriteLine(xml.ToString());
            xml.Save(Path.Combine(TempDir, "Out.xml"));
        }

        private static XElement ConvertDocToXml(WordprocessingDocument wDoc, int abstractNumId)
        {
            var xd = wDoc.MainDocumentPart.GetXDocument();
            // First, call RetrieveListItem so that all paragraphs are initialized with ListItemInfo
            var firstParagraph = xd.Descendants(W.p).FirstOrDefault();
            var listItem = ListItemRetriever.RetrieveListItem(wDoc, firstParagraph);
            var xml = new XElement("Root");
            var current = new Stack<XmlStackItem>();
            current.Push(new XmlStackItem { Element = xml, LevelNumbers = Array.Empty<int>() });
            foreach (var paragraph in xd.Descendants(W.p))
            {
                // The following does not take into account documents that have tracked revisions.
                // As necessary, call RevisionAccepter.AcceptRevisions before converting to XML.
                var text = paragraph.Descendants(W.t).Select(t => (string)t).StringConcatenate();
                var lii = paragraph.Annotation<ListItemRetriever.ListItemInfo>();
                if (lii.IsListItem && lii.AbstractNumId == abstractNumId)
                {
                    var levelNums = paragraph.Annotation<ListItemRetriever.LevelNumbers>();
                    if (levelNums.LevelNumbersArray.Length == current.Peek().LevelNumbers.Length)
                    {
                        current.Pop();
                        var levelNumsForThisIndent = levelNums.LevelNumbersArray;
                        var levelText = levelNums
                            .LevelNumbersArray.Select(l => l + ".")
                            .StringConcatenate()
                            .TrimEnd('.');
                        var newCurrentElement = new XElement("Indent", new XAttribute("Level", levelText));
                        current.Peek().Element.Add(newCurrentElement);
                        current.Push(
                            new XmlStackItem() { Element = newCurrentElement, LevelNumbers = levelNumsForThisIndent }
                        );
                        current.Peek().Element.Add(new XElement("Heading", text));
                    }
                    else if (levelNums.LevelNumbersArray.Length > current.Peek().LevelNumbers.Length)
                    {
                        for (var i = current.Peek().LevelNumbers.Length; i < levelNums.LevelNumbersArray.Length; i++)
                        {
                            var levelNumsForThisIndent = levelNums.LevelNumbersArray.Take(i + 1).ToArray();
                            var levelText = levelNums
                                .LevelNumbersArray.Select(l => l + ".")
                                .StringConcatenate()
                                .TrimEnd('.');
                            var newCurrentElement = new XElement("Indent", new XAttribute("Level", levelText));
                            current.Peek().Element.Add(newCurrentElement);
                            current.Push(
                                new XmlStackItem()
                                {
                                    Element = newCurrentElement,
                                    LevelNumbers = levelNumsForThisIndent,
                                }
                            );
                            current.Peek().Element.Add(new XElement("Heading", text));
                        }
                    }
                    else if (levelNums.LevelNumbersArray.Length < current.Peek().LevelNumbers.Length)
                    {
                        for (var i = current.Peek().LevelNumbers.Length; i > levelNums.LevelNumbersArray.Length; i--)
                            current.Pop();
                        current.Pop();
                        var levelNumsForThisIndent = levelNums.LevelNumbersArray;
                        var levelText = levelNums
                            .LevelNumbersArray.Select(l => l + ".")
                            .StringConcatenate()
                            .TrimEnd('.');
                        var newCurrentElement = new XElement("Indent", new XAttribute("Level", levelText));
                        current.Peek().Element.Add(newCurrentElement);
                        current.Push(
                            new XmlStackItem() { Element = newCurrentElement, LevelNumbers = levelNumsForThisIndent }
                        );
                        current.Peek().Element.Add(new XElement("Heading", text));
                    }
                }
                else
                {
                    current.Peek().Element.Add(new XElement("Paragraph", text));
                }
            }

            return xml;
        }
    }
}
