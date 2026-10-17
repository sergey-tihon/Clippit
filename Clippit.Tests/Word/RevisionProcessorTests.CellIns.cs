// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit.Word;

namespace Clippit.Tests.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
public partial class RevisionProcessorTests
{
    [Test]
    [Arguments(false, "A1 | B1 / A2 | B2")]
    [Arguments(true, "A1 | B1 / (span=2) A2")]
    public async Task RP061_RejectedInsertedCellWidensCellBefore(bool reject, string expectedRows)
    {
        var source = CreateDocument(
            """
            <w:tbl>
              <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
              <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
                <w:tc>
                  <w:tcPr>
                    <w:tcW w:w="0" w:type="auto"/>
                    <w:cellIns w:id="1" w:author="Author" w:date="2026-10-15T00:00:00Z"/>
                  </w:tcPr>
                  <w:p><w:r><w:t>B2</w:t></w:r></w:p>
                </w:tc>
              </w:tr>
            </w:tbl>
            <w:p/>
            """
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var table = doc.MainDocumentPart!.GetXDocument().Descendants(W.tbl).Single();
        await Assert.That(DescribeRows(table)).IsEqualTo(expectedRows);
    }
}
#pragma warning restore CA1707
