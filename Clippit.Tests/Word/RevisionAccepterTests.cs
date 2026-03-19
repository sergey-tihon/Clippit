// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word;

public class RevisionAccepterTests : TestsBase
{
    private static readonly XName[] s_trackedChangeElements =
    [
        W.ins,
        W.del,
        W.rPrChange,
        W.pPrChange,
        W.sectPrChange,
        W.tblPrChange,
        W.tcPrChange,
        W.trPrChange,
    ];

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

        // The output must contain no tracked-change markup anywhere in the main document part.
        var mainXDoc = doc.MainDocumentPart!.GetXDocument();
        var remaining = mainXDoc
            .Descendants()
            .Where(e => s_trackedChangeElements.Contains(e.Name))
            .Select(e => e.Name.LocalName)
            .Distinct()
            .ToList();

        await Assert.That(remaining).IsEmpty();
    }

    [Test]
    [Arguments("RA001-Tracked-Revisions-01.docx")]
    [Arguments("RA001-Tracked-Revisions-02.docx")]
    public async Task RA001_SourceDocuments_ContainTrackedChanges(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceDocx = new FileInfo(Path.Combine(sourceDir.FullName, name));

        using var doc = WordprocessingDocument.Open(sourceDocx.FullName, false);
        var mainXDoc = doc.MainDocumentPart!.GetXDocument();
        var trackedChangeCount = mainXDoc.Descendants().Count(e => s_trackedChangeElements.Contains(e.Name));

        // Verify the source documents actually have tracked changes to accept.
        await Assert.That(trackedChangeCount).IsGreaterThan(0);
    }
}
