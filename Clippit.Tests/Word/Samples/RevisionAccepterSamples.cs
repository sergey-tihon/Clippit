using Clippit.Word;
using Xunit;

namespace Clippit.Tests.Word.Samples
{
    public class RevisionAccepterSamples(ITestOutputHelper log) : TestsBase(log)
    {
        [Fact]
        public void Sample()
        {
            var srcDoc = new WmlDocument("../../../Word/Samples/RevisionAccepter/Source1.docx");
            var result = RevisionAccepter.AcceptRevisions(srcDoc);
            result.SaveAs(Path.Combine(TempDir, "Out1.docx"));
        }
    }
}
