using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Clippit.Word;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Word.Samples
{
    public class WmlComparerSamples : TestsBase
    {
        public WmlComparerSamples(ITestOutputHelper log) : base(log)
        {
        }
        
        private static string GetFilePath(string path) =>
            Path.Combine("../../../Word/Samples/WmlComparer/", path);

        [Fact]
        public void Sample1()
        {
            var settings = new WmlComparerSettings();
            var result = WmlComparer.Compare(
                new WmlDocument(GetFilePath("Sample1/Source1.docx")),
                new WmlDocument(GetFilePath("Sample1/Source2.docx")),
                settings);
            result.SaveAs(Path.Combine(TempDir, "Compared.docx"));

            var revisions = WmlComparer.GetRevisions(result, settings);
            foreach (var rev in revisions)
            {
                Log.WriteLine("Author: " + rev.Author);
                Log.WriteLine("Revision type: " + rev.RevisionType);
                Log.WriteLine("Revision text: " + rev.Text);
            }
        }

        [Fact]
        public void Sample2()
        {
            var originalWml = new WmlDocument(GetFilePath("Sample2/Original.docx"));
            var revisedDocumentInfoList = new List<WmlRevisedDocumentInfo>()
            {
                new()
                {
                    RevisedDocument = new WmlDocument(GetFilePath("Sample2/RevisedByBob.docx")),
                    Revisor = "Bob",
                    Color = Color.LightBlue,
                },
                new()
                {
                    RevisedDocument = new WmlDocument(GetFilePath("Sample2/RevisedByMary.docx")),
                    Revisor = "Mary",
                    Color = Color.LightYellow,
                },
            };
            
            var settings = new WmlComparerSettings();
            var consolidatedWml = WmlComparer.Consolidate(
                originalWml,
                revisedDocumentInfoList,
                settings);
            consolidatedWml.SaveAs(Path.Combine(TempDir, "Consolidated.docx"));
        }
    }
}
