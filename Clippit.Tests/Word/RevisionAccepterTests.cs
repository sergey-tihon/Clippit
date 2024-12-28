// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace Clippit.Tests.Word
{
    public class RevisionAccepterTests : TestsBase
    {
        public RevisionAccepterTests(ITestOutputHelper log)
            : base(log) { }

        [Theory]
        [InlineData("RA001-Tracked-Revisions-01.docx")]
        [InlineData("RA001-Tracked-Revisions-02.docx")]
        public void RA001(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

            var notAccepted = new WmlDocument(sourceDocx.FullName);
            var afterAccepting = RevisionAccepter.AcceptRevisions(notAccepted);
            var processedDestDocx = new FileInfo(
                Path.Combine(TempDir, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx"))
            );
            afterAccepting.SaveAs(processedDestDocx.FullName);
        }
    }
}

#endif
