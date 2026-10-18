// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit.Word;

namespace Clippit.Tests.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
public partial class RevisionProcessorTests
{
    // Run-level revisions wrapped in a run container belong to the enclosing paragraph.
    [Test]
    [Arguments("hyperlink", """w:anchor="Target" """, false)]
    [Arguments("hyperlink", """w:anchor="Target" """, true)]
    [Arguments("fldSimple", """w:instr=" AUTHOR " """, false)]
    [Arguments("fldSimple", """w:instr=" AUTHOR " """, true)]
    [Arguments("dir", """w:val="rtl" """, false)]
    [Arguments("dir", """w:val="rtl" """, true)]
    [Arguments("bdo", """w:val="rtl" """, false)]
    [Arguments("bdo", """w:val="rtl" """, true)]
    public async Task RP053_RevisionsInsideRunContainer(string container, string attributes, bool reject)
    {
        var source = CreateDocument(
            $"""
            <w:p>
              <w:{container} {attributes}>
                <w:r><w:t xml:space="preserve">Alpha </w:t></w:r>
                <w:del w:id="1" w:author="Author" w:date="2026-10-15T00:00:00Z">
                  <w:r><w:delText>beta</w:delText></w:r>
                </w:del>
                <w:ins w:id="2" w:author="Author" w:date="2026-10-15T00:00:00Z">
                  <w:r><w:t>delta</w:t></w:r>
                </w:ins>
                <w:r><w:t xml:space="preserve"> gamma</w:t></w:r>
              </w:{container}>
            </w:p>
            <w:p><w:bookmarkStart w:id="0" w:name="Target"/><w:r><w:t>Target</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>
            """
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc, RunContainerErrors(container));
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var wrapper = doc.MainDocumentPart!.GetXDocument().Descendants(W.w + container).Single();
        await Assert.That(TextOf(wrapper)).IsEqualTo(reject ? "Alpha beta gamma" : "Alpha delta gamma");
    }

    [Test]
    [Arguments("dir", false, "AlphaNext")]
    [Arguments("dir", true, "Alpha|Next")]
    [Arguments("bdo", false, "AlphaNext")]
    [Arguments("bdo", true, "Alpha|Next")]
    public async Task RP060_DeletedParagraphMarkAfterBidirectionalContainer(
        string container,
        bool reject,
        string expectedParagraphs
    )
    {
        var source = CreateDocument(
            $"""
            <w:p>
              <w:pPr><w:rPr><w:del w:id="1" w:author="Author" w:date="2026-10-15T00:00:00Z"/></w:rPr></w:pPr>
              <w:{container} w:val="rtl"><w:r><w:t>Alpha</w:t></w:r></w:{container}>
            </w:p>
            <w:p><w:r><w:t>Next</w:t></w:r></w:p>
            """
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc, RunContainerErrors(container));
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var paragraphs = doc.MainDocumentPart!.GetXDocument().Descendants(W.p).Select(TextOf);
        await Assert.That(string.Join("|", paragraphs)).IsEqualTo(expectedParagraphs);
    }

    // The Open XML SDK's validator does not know w:dir and w:bdo (ISO/IEC 29500 run containers that Word reads and
    // writes), so it reports them as unexpected children of w:p.
    private static List<string> RunContainerErrors(string container) =>
        container is "dir" or "bdo" ? [$"unexpected child element '{W.w.NamespaceName}:{container}'"] : [];
}
#pragma warning restore CA1707
