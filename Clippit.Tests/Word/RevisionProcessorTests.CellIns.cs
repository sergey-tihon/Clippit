// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Xml.Linq;
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

    [Test]
    [Arguments("cellIns", false, "Inner kept | Revised")]
    [Arguments("cellIns", true, "(span=2) Inner kept")]
    [Arguments("cellDel", false, "(span=2) Inner kept")]
    [Arguments("cellDel", true, "Inner kept | Revised")]
    public async Task RP067_NestedCellRevisionKeepsEnclosingCell(string marker, bool reject, string expectedRows)
    {
        var source = CreateDocument(
            $"""
            <w:tbl>
              <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>
              <w:tr><w:tc>
                <w:p><w:r><w:t>Outer before</w:t></w:r></w:p>
                <w:tbl>
                  <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
                  <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>Inner kept</w:t></w:r></w:p></w:tc>
                    <w:tc>
                      <w:tcPr><w:{marker} w:id="1" w:author="Author"/></w:tcPr>
                      <w:p><w:r><w:t>Revised</w:t></w:r></w:p>
                    </w:tc>
                  </w:tr>
                </w:tbl>
                <w:p><w:r><w:t>Outer after</w:t></w:r></w:p>
              </w:tc></w:tr>
            </w:tbl>
            <w:p/>
            """
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var outerCell = doc.MainDocumentPart!.GetXDocument()
            .Descendants(W.tbl)
            .First()
            .Element(W.tr)!
            .Elements(W.tc)
            .Single();
        await Assert
            .That(string.Join("|", outerCell.Elements(W.p).Select(TextOf)))
            .IsEqualTo("Outer before|Outer after");
        await Assert.That(DescribeRows(outerCell.Element(W.tbl)!)).IsEqualTo(expectedRows);
    }

    [Test]
    [Arguments(false, false, "Kept | Inserted")]
    [Arguments(true, false, "(span=2) Kept")]
    [Arguments(false, true, "Kept | Inserted")]
    [Arguments(true, true, "(span=2) Kept")]
    public async Task RP068_InsertedCellWithPropertyChangeIsRejected(
        bool reject,
        bool precedingPropertyChange,
        string expectedRows
    )
    {
        var precedingChange = precedingPropertyChange
            ? """
                <w:tcPrChange w:id="2" w:author="Author">
                  <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
                </w:tcPrChange>
                """
            : "";
        var source = CreateDocument(
            $"""
            <w:tbl>
              <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
              <w:tr>
                <w:tc><w:tcPr>{precedingChange}</w:tcPr><w:p><w:r><w:t>Kept</w:t></w:r></w:p></w:tc>
                <w:tc>
                  <w:tcPr>
                    <w:cellIns w:id="1" w:author="Author"/>
                    <w:tcPrChange w:id="3" w:author="Author"><w:tcPr/></w:tcPrChange>
                  </w:tcPr>
                  <w:p><w:r><w:t>Inserted</w:t></w:r></w:p>
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
        await Assert
            .That(DescribeRows(doc.MainDocumentPart!.GetXDocument().Descendants(W.tbl).Single()))
            .IsEqualTo(expectedRows);
    }

    [Test]
    [Arguments("cellIns", true, 1, 2, 1, "(span=3) Kept")]
    [Arguments("cellDel", false, 1, 2, 1, "(span=3) Kept")]
    [Arguments("cellIns", true, 2, 3, 2, "(span=8) Kept")]
    [Arguments("cellDel", false, 2, 3, 2, "(span=8) Kept")]
    public async Task RP069_RemovedCellsTransferAllGridColumns(
        string marker,
        bool reject,
        int precedingSpan,
        int removedSpan,
        int removedCount,
        string expectedRows
    )
    {
        var removedCells = string.Concat(
            Enumerable
                .Range(1, removedCount)
                .Select(id =>
                    $"""
                        <w:tc>
                          <w:tcPr><w:gridSpan w:val="{removedSpan}"/><w:{marker} w:id="{id}" w:author="Author"/></w:tcPr>
                          <w:p><w:r><w:t>Removed {id}</w:t></w:r></w:p>
                        </w:tc>
                        """
                )
        );
        var grid = string.Concat(
            Enumerable.Repeat("<w:gridCol w:w=\"1000\"/>", precedingSpan + removedSpan * removedCount)
        );
        var source = CreateDocument(
            $"""
            <w:tbl>
              <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
              <w:tblGrid>{grid}</w:tblGrid>
              <w:tr>
                <w:tc>
                  <w:tcPr><w:gridSpan w:val="{precedingSpan}"/></w:tcPr>
                  <w:p><w:r><w:t>Kept</w:t></w:r></w:p>
                </w:tc>
                {removedCells}
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
        await Assert
            .That(DescribeRows(doc.MainDocumentPart!.GetXDocument().Descendants(W.tbl).Single()))
            .IsEqualTo(expectedRows);
    }

    [Test]
    [Arguments("sdt", "inserted", true, false)]
    [Arguments("customXml", "inserted", true, false)]
    [Arguments("nested", "inserted", true, false)]
    [Arguments("sdt", "preceding", true, false)]
    [Arguments("customXml", "preceding", true, false)]
    [Arguments("nested", "preceding", true, false)]
    [Arguments("sdt", "separate", true, false)]
    [Arguments("customXml", "separate", true, false)]
    [Arguments("nested", "separate", true, false)]
    [Arguments("sdt", "separate", true, true)]
    [Arguments("customXml", "separate", true, true)]
    [Arguments("nested", "separate", true, true)]
    [Arguments("sdt", "separate", false, false)]
    [Arguments("customXml", "separate", false, false)]
    [Arguments("nested", "separate", false, false)]
    [Arguments("sdt", "together", true, false)]
    [Arguments("customXml", "together", true, false)]
    [Arguments("nested", "together", true, false)]
    public async Task RP070_WrappedCellRevisionPreservesStructure(
        string wrapper,
        string layout,
        bool reject,
        bool precedingPropertyChange
    )
    {
        string Wrap(string cells, string label)
        {
            var customXml = $"<w:customXml w:element=\"{label}\"><w:customXmlPr/>{cells}</w:customXml>";
            return wrapper switch
            {
                "sdt" =>
                    $"<w:sdt><w:sdtPr><w:tag w:val=\"{label}\"/></w:sdtPr><w:sdtContent>{cells}</w:sdtContent></w:sdt>",
                "customXml" => customXml,
                "nested" =>
                    $"<w:sdt><w:sdtPr><w:tag w:val=\"{label}\"/></w:sdtPr><w:sdtContent>{customXml}</w:sdtContent></w:sdt>",
                _ => throw new ArgumentOutOfRangeException(nameof(wrapper)),
            };
        }

        var propertyChange = precedingPropertyChange
            ? """
                <w:tcPrChange w:id="2" w:author="Author">
                  <w:tcPr><w:gridSpan w:val="3"/><w:vAlign w:val="center"/></w:tcPr>
                </w:tcPrChange>
                """
            : "";
        var precedingCell = $"""
            <w:tc>
              <w:tcPr><w:tcW w:w="0" w:type="auto"/><w:vAlign w:val="center"/>{propertyChange}</w:tcPr>
              <w:p><w:r><w:t>Kept</w:t></w:r></w:p>
            </w:tc>
            """;
        var marker = reject ? "cellIns" : "cellDel";
        var revisedCell = $"""
            <w:tc>
              <w:tcPr><w:gridSpan w:val="2"/><w:{marker} w:id="1" w:author="Author"/></w:tcPr>
              <w:p><w:r><w:t>Removed</w:t></w:r></w:p>
            </w:tc>
            """;
        const string tailCell = "<w:tc><w:p><w:r><w:t>Tail</w:t></w:r></w:p></w:tc>";
        var cells = layout switch
        {
            "inserted" => precedingCell + Wrap(revisedCell, "removed") + tailCell,
            "preceding" => Wrap(precedingCell, "kept") + revisedCell + tailCell,
            "separate" => Wrap(precedingCell, "kept") + Wrap(revisedCell, "removed") + tailCell,
            "together" => Wrap(precedingCell + revisedCell + tailCell, "group"),
            _ => throw new ArgumentOutOfRangeException(nameof(layout)),
        };
        var source = CreateDocument(
            $"""
            <w:tbl>
              <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>
              <w:tr>{cells}</w:tr>
            </w:tbl><w:p/>
            """
        );
        using (var sourceStream = new OpenXmlMemoryStreamDocument(source))
        using (var sourceDoc = sourceStream.GetWordprocessingDocument())
        {
            await Validate(sourceDoc);
        }

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var row = doc.MainDocumentPart!.GetXDocument().Descendants(W.tr).Single();
        var remainingCells = row.Descendants(W.tc).ToArray();
        await Assert.That(string.Join("|", remainingCells.Select(TextOf))).IsEqualTo("Kept|Tail");
        var kept = remainingCells[0];
        await Assert.That((int?)kept.Element(W.tcPr)?.Element(W.gridSpan)?.Attribute(W.val)).IsEqualTo(3);
        await Assert.That((string?)kept.Element(W.tcPr)?.Element(W.vAlign)?.Attribute(W.val)).IsEqualTo("center");
        var expectedPath =
            layout == "inserted"
                ? ""
                : wrapper switch
                {
                    "sdt" => "sdtContent/sdt",
                    "customXml" => "customXml",
                    _ => "customXml/sdtContent/sdt",
                };
        await Assert
            .That(string.Join("/", kept.Ancestors().TakeWhile(e => e != row).Select(e => e.Name.LocalName)))
            .IsEqualTo(expectedPath);
        var labels = layout switch
        {
            "inserted" => "removed",
            "preceding" => "kept",
            "together" => "group",
            _ => "kept|removed",
        };
        await Assert
            .That(string.Join("|", row.Descendants(W.tag).Attributes(W.val).Select(a => a.Value)))
            .IsEqualTo(wrapper == "customXml" ? "" : labels);
        await Assert
            .That(string.Join("|", row.Descendants(W.customXml).Attributes(W.element).Select(a => a.Value)))
            .IsEqualTo(wrapper == "sdt" ? "" : labels);
    }
}
#pragma warning restore CA1707
