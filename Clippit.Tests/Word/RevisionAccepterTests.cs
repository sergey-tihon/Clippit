// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word;

public class RevisionAccepterTests : TestsBase
{
    [Test]
    [Arguments("RA001-Tracked-Revisions-01.docx")]
    [Arguments("RA001-Tracked-Revisions-02.docx")]
    public async Task RA001_AcceptRevisions_RemovesAllTrackedChanges(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

        var notAccepted = new WmlDocument(sourceDocx.FullName);
        var afterAccepting = RevisionAccepter.AcceptRevisions(notAccepted);
        var processedDestDocx = new FileInfo(
            Path.Combine(TempDir, sourceDocx.Name.Replace(".docx", "-processed-by-RevisionAccepter.docx"))
        );
        afterAccepting.SaveAs(processedDestDocx.FullName);

        // Verify the output is a well-formed Word document that can be opened.
        using var doc = WordprocessingDocument.Open(processedDestDocx.FullName, false);

        // The output must contain no tracked-change markup anywhere in the main document part,
        // using the library's own canonical definition of tracked revisions.
        var hasTrackedRevisions = RevisionAccepter.PartHasTrackedRevisions(doc.MainDocumentPart!);

        await Assert.That(hasTrackedRevisions).IsFalse();
    }

    [Test]
    [Arguments("RA001-Tracked-Revisions-01.docx")]
    [Arguments("RA001-Tracked-Revisions-02.docx")]
    public async Task RA001_SourceDocuments_ContainTrackedChanges(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

        using var doc = WordprocessingDocument.Open(sourceDocx.FullName, false);
        var hasTrackedRevisions = RevisionAccepter.HasTrackedRevisions(doc);

        // Verify the source documents actually have tracked changes to accept.
        await Assert.That(hasTrackedRevisions).IsTrue();
    }
}
