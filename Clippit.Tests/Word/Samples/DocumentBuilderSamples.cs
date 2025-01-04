using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Clippit.Tests.Word.Samples
{
    public class DocumentBuilderSamples : TestsBase
    {
        public DocumentBuilderSamples(ITestOutputHelper log)
            : base(log) { }

        private static string GetFilePath(string path) => Path.Combine("../../../Word/Samples/DocumentBuilder/", path);

        [Fact]
        public void Sample1()
        {
            var source1 = GetFilePath("Sample1/Source1.docx");
            var source2 = GetFilePath("Sample1/Source2.docx");
            var source3 = GetFilePath("Sample1/Source3.docx");

            // Create new document from 10 paragraphs starting at paragraph 5 of Source1.docx
            var sources = new List<ISource> { new Source(new WmlDocument(source1), 5, 10, true) };
            DocumentBuilder.BuildDocument(sources, Path.Combine(TempDir, "Out1.docx"));

            // Create new document from paragraph 1, and paragraphs 5 through end of Source3.docx.
            // This effectively 'deletes' paragraphs 2-4
            sources = new List<ISource>()
            {
                new Source(new WmlDocument(source3), 0, 1, false),
                new Source(new WmlDocument(source3), 4, false),
            };
            DocumentBuilder.BuildDocument(sources, Path.Combine(TempDir, "Out2.docx"));

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source1.
            sources = new List<ISource>()
            {
                new Source(new WmlDocument(source1), true),
                new Source(new WmlDocument(source2), false),
            };
            DocumentBuilder.BuildDocument(sources, Path.Combine(TempDir, "Out3.docx"));

            // Create a new document that consists of the entirety of Source1.docx and Source2.docx.  Use
            // the section information (headings and footers) from source2.
            sources = new List<ISource>()
            {
                new Source(new WmlDocument(source1), false),
                new Source(new WmlDocument(source2), true),
            };
            DocumentBuilder.BuildDocument(sources, Path.Combine(TempDir, "Out4.docx"));

            // Create a new document that consists of the first 5 paragraphs of Source1.docx and the first
            // five paragraphs of Source2.docx.  This example returns a new WmlDocument, when you then can
            // serialize to a SharePoint document library, or use in some other interesting scenario.
            sources = new List<ISource>()
            {
                new Source(new WmlDocument(source1), 0, 5, false),
                new Source(new WmlDocument(source2), 0, 5, true),
            };
            var out5 = DocumentBuilder.BuildDocument(sources);
            out5.SaveAs(Path.Combine(TempDir, "Out5.docx")); // save it to the file system, but we could just as easily done something
            // else with it.
        }

        private class DocumentInfo
        {
            public int DocumentNumber { get; init; }
            public int Start { get; init; }
            public int Count { get; init; }
        }

        [Fact]
        public void Sample2()
        {
            // Insert an abstract and author biography into a white paper.
            var sources = new List<ISource>
            {
                new Source(new WmlDocument(GetFilePath("Sample2/WhitePaper.docx")), 0, 1, true),
                new Source(new WmlDocument(GetFilePath("Sample2/Abstract.docx")), false),
                new Source(new WmlDocument(GetFilePath("Sample2/AuthorBiography.docx")), false),
                new Source(new WmlDocument(GetFilePath("Sample2/WhitePaper.docx")), 1, false),
            };
            DocumentBuilder.BuildDocument(sources, Path.Combine(TempDir, "AssembledPaper.docx"));

            // Delete all paragraphs with a specific style.
            using (var doc = WordprocessingDocument.Open(GetFilePath("Sample2/Notes.docx"), false))
            {
                sources = doc
                    .MainDocumentPart.GetXDocument()
                    .Root.Element(W.body)
                    .Elements()
                    .Select((p, i) => new { Paragraph = p, Index = i })
                    .GroupAdjacent(pi =>
                        (string)pi.Paragraph.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault()
                        != "Note"
                    )
                    .Where(g => g.Key)
                    .Select(g => new Source(
                        new WmlDocument(GetFilePath("Sample2/Notes.docx")),
                        g.First().Index,
                        g.Last().Index - g.First().Index + 1,
                        true
                    ))
                    .Cast<ISource>()
                    .ToList();
            }

            DocumentBuilder.BuildDocument(sources, Path.Combine(TempDir, "NewNotes.docx"));

            // Shred a document into multiple parts for each section
            List<DocumentInfo> documentList;
            using (var doc = WordprocessingDocument.Open(GetFilePath("Sample2/Spec.docx"), false))
            {
                var sectionCounts = doc
                    .MainDocumentPart.GetXDocument()
                    .Root.Element(W.body)
                    .Elements()
                    .Rollup(
                        0,
                        (pi, last) =>
                            (string)pi.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault()
                            == "Heading1"
                                ? last + 1
                                : last
                    );
                var beforeZipped = doc
                    .MainDocumentPart.GetXDocument()
                    .Root.Element(W.body)
                    .Elements()
                    .Select((p, i) => new { Paragraph = p, Index = i });
                var zipped = PtExtensions.PtZip(
                    beforeZipped,
                    sectionCounts,
                    (pi, sc) =>
                        new
                        {
                            Paragraph = pi.Paragraph,
                            Index = pi.Index,
                            SectionIndex = sc,
                        }
                );
                documentList = zipped
                    .GroupAdjacent(p => p.SectionIndex)
                    .Select(g => new DocumentInfo
                    {
                        DocumentNumber = g.Key,
                        Start = g.First().Index,
                        Count = g.Last().Index - g.First().Index + 1,
                    })
                    .ToList();
            }

            foreach (var doc in documentList)
            {
                var fileName = $"Section{doc.DocumentNumber:000}.docx";
                var documentSource = new List<ISource>
                {
                    new Source(new WmlDocument(GetFilePath("Sample2/Spec.docx")), doc.Start, doc.Count, true),
                };
                DocumentBuilder.BuildDocument(documentSource, Path.Combine(TempDir, fileName));
            }

            // Re-assemble the parts into a single document.
            sources = new DirectoryInfo(TempDir)
                .GetFiles("Section*.docx")
                .Select(d => new Source(new WmlDocument(d.FullName), true))
                .Cast<ISource>()
                .ToList();
            DocumentBuilder.BuildDocument(sources, Path.Combine(TempDir, "ReassembledSpec.docx"));
        }

        [Fact]
        public void Sample3()
        {
            var doc1 = new WmlDocument(GetFilePath("Sample3/Template.docx"));
            using (var mem = new MemoryStream())
            {
                mem.Write(doc1.DocumentByteArray, 0, doc1.DocumentByteArray.Length);
                using (var doc = WordprocessingDocument.Open(mem, true))
                {
                    var xDoc = doc.MainDocumentPart.GetXDocument();
                    var frontMatterPara = xDoc.Root.Descendants(W.txbxContent).Elements(W.p).FirstOrDefault();
                    frontMatterPara.ReplaceWith(new XElement(PtOpenXml.Insert, new XAttribute("Id", "Front")));
                    var tbl = xDoc.Root.Element(W.body).Elements(W.tbl).FirstOrDefault();
                    var firstCell = tbl.Descendants(W.tr).First().Descendants(W.p).First();
                    firstCell.ReplaceWith(new XElement(PtOpenXml.Insert, new XAttribute("Id", "Liz")));
                    var secondCell = tbl.Descendants(W.tr).Skip(1).First().Descendants(W.p).First();
                    secondCell.ReplaceWith(new XElement(PtOpenXml.Insert, new XAttribute("Id", "Eric")));
                    doc.MainDocumentPart.PutXDocument();
                }
                doc1.DocumentByteArray = mem.ToArray();
            }

            var outFileName = Path.Combine(TempDir, "Out.docx");
            var sources = new List<ISource>()
            {
                new Source(doc1, true),
                new Source(new WmlDocument(GetFilePath("Sample3/Insert-01.docx")), "Liz"),
                new Source(new WmlDocument(GetFilePath("Sample3/Insert-02.docx")), "Eric"),
                new Source(new WmlDocument(GetFilePath("Sample3/FrontMatter.docx")), "Front"),
            };
            DocumentBuilder.BuildDocument(sources, outFileName);
        }

        [Fact]
        public void Sample4()
        {
            var solarSystemDoc = new WmlDocument(GetFilePath("Sample4/solar-system.docx"));
            using var streamDoc = new OpenXmlMemoryStreamDocument(solarSystemDoc);
            using var solarSystem = streamDoc.GetWordprocessingDocument();
            // get children elements of the <w:body> element
            var q1 = solarSystem.MainDocumentPart.GetXDocument().Root.Element(W.body).Elements();

            // project collection of tuples containing element and type
            var q2 = q1.Select(e =>
                {
                    var keyForGroupAdjacent = ".NonContentControl";
                    if (e.Name == W.sdt)
                        keyForGroupAdjacent = e.Element(W.sdtPr).Element(W.tag).Attribute(W.val).Value;
                    if (e.Name == W.sectPr)
                        keyForGroupAdjacent = null;
                    return new { Element = e, KeyForGroupAdjacent = keyForGroupAdjacent };
                })
                .Where(e => e.KeyForGroupAdjacent != null);

            // group by type
            var q3 = q2.GroupAdjacent(e => e.KeyForGroupAdjacent);

            // temporary code to dump q3
            foreach (var g in q3)
                Log.WriteLine("{0}:  {1}", g.Key, g.Count());
            //Environment.Exit(0);


            // validate existence of files referenced in content controls
            foreach (var f in q3.Where(g => g.Key != ".NonContentControl"))
            {
                var filename = GetFilePath("Sample4/" + f.Key + ".docx");
                var fi = new FileInfo(filename);
                if (!fi.Exists)
                {
                    Log.WriteLine("{0} doesn't exist.", filename);
                    Environment.Exit(0);
                }
            }

            // project collection with opened WordProcessingDocument
            var q4 = q3.Select(g => new
            {
                Group = g,
                Document = g.Key != ".NonContentControl"
                    ? new WmlDocument(GetFilePath("Sample4/" + g.Key + ".docx"))
                    : solarSystemDoc,
            });

            // project collection of OpenXml.PowerTools.Source
            var sources = q4.Select(g =>
                {
                    if (g.Group.Key == ".NonContentControl")
                        return new Source(
                            g.Document,
                            g.Group.First().Element.ElementsBeforeSelf().Count(),
                            g.Group.Count(),
                            false
                        );
                    else
                        return new Source(g.Document, false);
                })
                .Cast<ISource>()
                .ToList();

            DocumentBuilder.BuildDocument(sources, Path.Combine(TempDir, "solar-system-new.docx"));
        }
    }
}
