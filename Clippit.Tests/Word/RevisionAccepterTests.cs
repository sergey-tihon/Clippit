// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

public class RevisionAccepterTests : TestsBase
{
    [Test]
    [Arguments("RA001-Tracked-Revisions-01.docx")]
    [Arguments("RA001-Tracked-Revisions-02.docx")]
    [Arguments("RA001-Tracked-Revisions-02.docx")]
    public async Task RA001(string name)
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
