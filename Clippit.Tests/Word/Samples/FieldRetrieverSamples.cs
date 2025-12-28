using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word.Samples
{
    public class FieldRetrieverSamples() : Clippit.Tests.TestsBase
    {
        private static string GetFilePath(string path) => Path.Combine("../../../Word/Samples/FieldRetriever/", path);

        [Test]
        public void Sample1()
        {
            var docWithFooter = new FileInfo(GetFilePath("DocWithFooter1.docx"));
            var scrubbedDocument = new FileInfo(Path.Combine(TempDir, "DocWithFooterScrubbed1.docx"));
            File.Copy(docWithFooter.FullName, scrubbedDocument.FullName, true);
            using var wDoc = WordprocessingDocument.Open(scrubbedDocument.FullName, true);
            ScrubFooter(wDoc, ["PAGE"]);
        }

        [Test]
        public void Sample2()
        {
            var docWithFooter = new FileInfo(GetFilePath("DocWithFooter2.docx"));
            var scrubbedDocument = new FileInfo(Path.Combine(TempDir, "DocWithFooterScrubbed2.docx"));
            File.Copy(docWithFooter.FullName, scrubbedDocument.FullName, true);
            using var wDoc = WordprocessingDocument.Open(scrubbedDocument.FullName, true);
            ScrubFooter(wDoc, ["PAGE", "DATE"]);
        }

        private static void ScrubFooter(WordprocessingDocument wDoc, string[] fieldTypesToKeep)
        {
            foreach (var footer in wDoc.MainDocumentPart.FooterParts)
            {
                FieldRetriever.AnnotateWithFieldInfo(footer);
                var root = footer.GetXDocument().Root;
                RemoveAllButSpecificFields(root, fieldTypesToKeep);
                footer.PutXDocument();
            }
        }

        private static void RemoveAllButSpecificFields(XElement root, string[] fieldTypesToRetain)
        {
            var cachedAnnotationInformation = root.Annotation<Dictionary<int, List<XElement>>>();
            var runsToKeep = cachedAnnotationInformation
                .SelectMany(item =>
                    root.Descendants()
                        .Where(d =>
                        {
                            var stack = d.Annotation<Stack<FieldRetriever.FieldElementTypeInfo>>();
                            return stack != null && stack.Any(stackItem => stackItem.Id == item.Key);
                        })
                        .Select(d => d.AncestorsAndSelf(W.r).FirstOrDefault())
                        .GroupAdjacent(o => o)
                        .Select(g => g.First())
                        .ToList()
                )
                .ToList();
            foreach (var paragraph in root.Descendants(W.p).ToList())
            {
                if (paragraph.Elements(W.r).Any(r => runsToKeep.Contains(r)))
                {
                    paragraph.Elements(W.r).Where(r => !runsToKeep.Contains(r) && !r.Elements(W.tab).Any()).Remove();
                    paragraph
                        .Elements(W.r)
                        .Where(r => !runsToKeep.Contains(r))
                        .Elements()
                        .Where(rc => rc.Name != W.rPr && rc.Name != W.tab)
                        .Remove();
                }
                else
                {
                    paragraph.Remove();
                }
            }

            root.Descendants(W.tbl).Remove();
        }
    }
}
