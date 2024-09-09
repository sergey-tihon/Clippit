using System.IO;
using Clippit.Word;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Word.Samples
{
    public class RevisionAccepterSamples : TestsBase
    {
        public RevisionAccepterSamples(ITestOutputHelper log)
            : base(log) { }

        [Fact]
        public void Sample()
        {
            var srcDoc = new WmlDocument("../../../Word/Samples/RevisionAccepter/Source1.docx");
            var result = RevisionAccepter.AcceptRevisions(srcDoc);
            result.SaveAs(Path.Combine(TempDir, "Out1.docx"));
        }
    }
}
