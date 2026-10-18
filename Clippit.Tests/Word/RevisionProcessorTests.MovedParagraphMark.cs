// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit.Word;

namespace Clippit.Tests.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
public partial class RevisionProcessorTests
{
    [Test]
    [Arguments(false, "Before|Middle|Moved|After")]
    [Arguments(true, "Before|Moved|Middle|After")]
    public async Task RP056_MovedParagraphLeavesNoEmptyParagraph(bool reject, string expectedParagraphs)
    {
        var source = CreateDocument(
            """
            <w:p><w:r><w:t>Before</w:t></w:r></w:p>
            <w:p>
              <w:pPr><w:rPr><w:moveFrom w:id="1" w:author="Author" w:date="2026-10-15T00:00:00Z"/></w:rPr></w:pPr>
              <w:moveFromRangeStart w:id="2" w:author="Author" w:date="2026-10-15T00:00:00Z" w:name="move1"/>
              <w:moveFrom w:id="3" w:author="Author" w:date="2026-10-15T00:00:00Z">
                <w:r><w:t>Moved</w:t></w:r>
              </w:moveFrom>
            </w:p>
            <w:moveFromRangeEnd w:id="2"/>
            <w:p><w:r><w:t>Middle</w:t></w:r></w:p>
            <w:p>
              <w:pPr><w:rPr><w:moveTo w:id="4" w:author="Author" w:date="2026-10-15T00:00:00Z"/></w:rPr></w:pPr>
              <w:moveToRangeStart w:id="5" w:author="Author" w:date="2026-10-15T00:00:00Z" w:name="move1"/>
              <w:moveTo w:id="6" w:author="Author" w:date="2026-10-15T00:00:00Z">
                <w:r><w:t>Moved</w:t></w:r>
              </w:moveTo>
            </w:p>
            <w:moveToRangeEnd w:id="5"/>
            <w:p><w:r><w:t>After</w:t></w:r></w:p>
            """
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var paragraphs = doc.MainDocumentPart!.GetXDocument().Descendants(W.p).Select(TextOf);
        await Assert.That(string.Join("|", paragraphs)).IsEqualTo(expectedParagraphs);
    }
}
#pragma warning restore CA1707
