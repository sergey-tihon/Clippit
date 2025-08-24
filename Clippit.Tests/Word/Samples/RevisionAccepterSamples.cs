using Clippit.Word;

namespace Clippit.Tests.Word.Samples
{
    public class RevisionAccepterSamples() : Clippit.Tests.TestsBase
    {
        [Test]
        public void Sample()
        {
            var srcDoc = new WmlDocument("../../../Word/Samples/RevisionAccepter/Source1.docx");
            var result = RevisionAccepter.AcceptRevisions(srcDoc);
            result.SaveAs(Path.Combine(TempDir, "Out1.docx"));
        }
    }
}