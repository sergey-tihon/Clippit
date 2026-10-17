// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
public partial class RevisionProcessorTests : TestsBase
{
    [Test]
    [Arguments("RP/RP002-Deleted-Text.docx")]
    [Arguments("RP/RP003-Inserted-Text.docx")]
    [Arguments("RP/RP004-Deleted-Text-in-CC.docx")]
    [Arguments("RP/RP005-Deleted-Paragraph-Mark.docx")]
    [Arguments("RP/RP006-Inserted-Paragraph-Mark.docx")]
    [Arguments("RP/RP007-Multiple-Deleted-Para-Mark.docx")]
    [Arguments("RP/RP008-Multiple-Inserted-Para-Mark.docx")]
    [Arguments("RP/RP009-Deleted-Table-Row.docx")]
    [Arguments("RP/RP010-Inserted-Table-Row.docx")]
    [Arguments("RP/RP011-Multiple-Deleted-Rows.docx")]
    [Arguments("RP/RP012-Multiple-Inserted-Rows.docx")]
    [Arguments("RP/RP013-Deleted-Math-Control-Char.docx")]
    [Arguments("RP/RP014-Inserted-Math-Control-Char.docx")]
    [Arguments("RP/RP015-MoveFrom-MoveTo.docx")]
    [Arguments("RP/RP016-Deleted-CC.docx")]
    [Arguments("RP/RP017-Inserted-CC.docx")]
    [Arguments("RP/RP018-MoveFrom-MoveTo-CC.docx")]
    [Arguments("RP/RP019-Deleted-Field-Code.docx")]
    [Arguments("RP/RP020-Inserted-Field-Code.docx")]
    [Arguments("RP/RP021-Inserted-Numbering-Properties.docx")]
    [Arguments("RP/RP022-NumberingChange.docx")]
    [Arguments("RP/RP023-NumberingChange.docx")]
    [Arguments("RP/RP024-ParagraphMark-rPr-Change.docx")]
    [Arguments("RP/RP025-Paragraph-Props-Change.docx")]
    [Arguments("RP/RP026-NumberingChange.docx")]
    [Arguments("RP/RP027-Change-Section.docx")]
    [Arguments("RP/RP028-Table-Grid-Change.docx")]
    [Arguments("RP/RP029-Table-Row-Props-Change.docx")]
    [Arguments("RP/RP030-Table-Row-Props-Change.docx")]
    [Arguments("RP/RP031-Table-Prop-Change.docx")]
    [Arguments("RP/RP032-Table-Prop-Change.docx")]
    [Arguments("RP/RP033-Table-Prop-Ex-Change.docx")]
    [Arguments("RP/RP034-Deleted-Cells.docx")]
    [Arguments("RP/RP035-Inserted-Cells.docx")]
    [Arguments("RP/RP036-Vert-Merged-Cells.docx")]
    [Arguments("RP/RP037-Changed-Style-Para-Props.docx")]
    [Arguments("RP/RP038-Inserted-Paras-at-End.docx")]
    [Arguments("RP/RP039-Inserted-Paras-at-End.docx")]
    [Arguments("RP/RP040-Deleted-Paras-at-End.docx")]
    [Arguments("RP/RP041-Cell-With-Empty-Paras-at-End.docx")]
    [Arguments("RP/RP042-Deleted-Para-Mark-at-End.docx")]
    [Arguments("RP/RP043-MERGEFORMAT-Field-Code.docx")]
    [Arguments("RP/RP044-MERGEFORMAT-Field-Code.docx")]
    [Arguments("RP/RP045-One-and-Half-Deleted-Lines-at-End.docx")]
    [Arguments("RP/RP046-Consecutive-Deleted-Ranges.docx")]
    [Arguments("RP/RP047-Inserted-and-Deleted-Paragraph-Mark.docx")]
    [Arguments("RP/RP048-Deleted-Inserted-Para-Mark.docx")]
    [Arguments("RP/RP049-Deleted-Para-Before-Table.docx")]
    [Arguments("RP/RP050-Deleted-Footnote.docx")]
    [Arguments("RP/RP052-Deleted-Para-Mark.docx")]
    public void RP001(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceFi = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var baselineAcceptedFi = new FileInfo(
            Path.Combine(sourceDir.FullName, name.Replace(".docx", "-Accepted.docx"))
        );
        var baselineRejectedFi = new FileInfo(
            Path.Combine(sourceDir.FullName, name.Replace(".docx", "-Rejected.docx"))
        );
        var sourceWml = new WmlDocument(sourceFi.FullName);
        var afterAcceptingWml = RevisionProcessor.AcceptRevisions(sourceWml);
        var afterRejectingWml = RevisionProcessor.RejectRevisions(sourceWml);

        var processedAcceptedFi = new FileInfo(Path.Combine(TempDir, sourceFi.Name.Replace(".docx", "-Accepted.docx")));
        afterAcceptingWml.SaveAs(processedAcceptedFi.FullName);

        var processedRejectedFi = new FileInfo(Path.Combine(TempDir, sourceFi.Name.Replace(".docx", "-Rejected.docx")));
        afterRejectingWml.SaveAs(processedRejectedFi.FullName);

        var settings = new WmlComparerSettings();

        if (!baselineAcceptedFi.Exists)
            Assert.Fail($"No Accepted baseline document for {name}");
        var acceptedResult = WmlComparer.Compare(
            new WmlDocument(baselineAcceptedFi.FullName),
            afterAcceptingWml,
            settings
        );
        var acceptedRevisions = WmlComparer.GetRevisions(acceptedResult, settings);
        if (acceptedRevisions.Any())
            Assert.Fail($"Regression: Accepted output differs from baseline for {name}");

        if (!baselineRejectedFi.Exists)
            Assert.Fail($"No Rejected baseline document for {name}");
        var rejectedResult = WmlComparer.Compare(
            new WmlDocument(baselineRejectedFi.FullName),
            afterRejectingWml,
            settings
        );
        var rejectedRevisions = WmlComparer.GetRevisions(rejectedResult, settings);
        if (rejectedRevisions.Any())
            Assert.Fail($"Regression: Rejected output differs from baseline for {name}");
    }

    [Test]
    [Arguments(false, "A1 | B1 / (span=2) A2")]
    [Arguments(true, "A1 | B1 / A2 | B2")]
    public async Task RP062_DeletedCellAfterCellWithoutProperties(bool reject, string expectedRows)
    {
        var source = CreateDocument(
            """
            <w:tbl>
              <w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
              <w:tr>
                <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
              </w:tr>
              <w:tr>
                <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
                <w:tc>
                  <w:tcPr>
                    <w:tcW w:w="0" w:type="auto"/>
                    <w:cellDel w:id="1" w:author="Author" w:date="2026-10-15T00:00:00Z"/>
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

    private static string DescribeRows(XElement table) =>
        string.Join(
            " / ",
            table
                .Elements(W.tr)
                .Select(tr =>
                    string.Join(
                        " | ",
                        tr.Elements(W.tc)
                            .Select(tc =>
                                (string?)tc.Elements(W.tcPr).Elements(W.gridSpan).Attributes(W.val).FirstOrDefault()
                                    is { } span
                                    ? $"(span={span}) {TextOf(tc)}"
                                    : TextOf(tc)
                            )
                    )
                )
        );

    private static string TextOf(XElement element) => string.Concat(element.Descendants(W.t).Select(t => t.Value));

    // With sectionChanged the section break also carries a w:sectPrChange; it then has no header or footer
    // references, whose fate under a w:sectPrChange RP055 covers.
    [Test]
    [Arguments(false, false, "center", "rIdHeader1 rIdFooter1 pgSz pgMar@1440 | rIdHeader2 rIdFooter2 pgSz pgMar@1440")]
    [Arguments(true, false, null, "rIdHeader1 rIdFooter1 pgSz pgMar@1440 | rIdHeader2 rIdFooter2 pgSz pgMar@1440")]
    [Arguments(false, true, "center", "pgSz pgMar@720 | rIdHeader2 rIdFooter2 pgSz pgMar@1440")]
    [Arguments(true, true, null, "pgSz pgMar@1440 | rIdHeader2 rIdFooter2 pgSz pgMar@1440")]
    public async Task RP058_ParagraphPropertiesChangeKeepsSectionBreak(
        bool reject,
        bool sectionChanged,
        string? expectedJustification,
        string expectedSections
    )
    {
        const string margins =
            """w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0" """;
        var sectionBreak = sectionChanged
            ? $"""
                <w:sectPr>
                  <w:pgSz w:w="11906" w:h="16838"/>
                  <w:pgMar w:top="720" {margins}/>
                  <w:sectPrChange w:id="1" w:author="Author" w:date="2026-10-15T00:00:00Z">
                    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" {margins}/></w:sectPr>
                  </w:sectPrChange>
                </w:sectPr>
                """
            : $"""
                <w:sectPr>
                  <w:headerReference w:type="default" r:id="rIdHeader1"/>
                  <w:footerReference w:type="default" r:id="rIdFooter1"/>
                  <w:pgSz w:w="11906" w:h="16838"/>
                  <w:pgMar w:top="1440" {margins}/>
                </w:sectPr>
                """;
        var source = CreateDocument(
            $"""
            <w:p>
              <w:pPr>
                <w:jc w:val="center"/>
                {sectionBreak}
                <w:pPrChange w:id="2" w:author="Author" w:date="2026-10-15T00:00:00Z"><w:pPr/></w:pPrChange>
              </w:pPr>
              <w:r><w:t>First section</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>Second section</w:t></w:r></w:p>
            <w:sectPr>
              <w:headerReference w:type="default" r:id="rIdHeader2"/>
              <w:footerReference w:type="default" r:id="rIdFooter2"/>
              <w:pgSz w:w="11906" w:h="16838"/>
              <w:pgMar w:top="1440" {margins}/>
            </w:sectPr>
            """,
            mainPart =>
            {
                foreach (var n in new[] { 1, 2 })
                {
                    mainPart
                        .AddNewPart<HeaderPart>($"rIdHeader{n}")
                        .PutXDocument(
                            XDocument.Parse(
                                $"<w:hdr {s_namespaces}><w:p><w:r><w:t>Header {n}</w:t></w:r></w:p></w:hdr>"
                            )
                        );
                    mainPart
                        .AddNewPart<FooterPart>($"rIdFooter{n}")
                        .PutXDocument(
                            XDocument.Parse(
                                $"<w:ftr {s_namespaces}><w:p><w:r><w:t>Footer {n}</w:t></w:r></w:p></w:ftr>"
                            )
                        );
                }
            }
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var body = doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!;
        var justification = (string?)
            body.Elements(W.p).First().Elements(W.pPr).Elements(W.jc).Attributes(W.val).FirstOrDefault();
        await Assert.That(justification).IsEqualTo(expectedJustification);
        var sections = body.Descendants(W.sectPr).Select(DescribeSection);
        await Assert.That(string.Join(" | ", sections)).IsEqualTo(expectedSections);
    }

    [Test]
    [Arguments(false, 720)]
    [Arguments(true, 1440)]
    public async Task RP055_SectionPropertiesChangeKeepsHeadersAndFooters(bool reject, int expectedTopMargin)
    {
        var source = CreateDocument(
            $"""
            <w:p>
              <w:pPr>{TrackedSectionProperties(1)}</w:pPr>
              <w:r><w:t>First section</w:t></w:r>
            </w:p>
            <w:p><w:r><w:t>Second section</w:t></w:r></w:p>
            {TrackedSectionProperties(2)}
            """,
            mainPart =>
            {
                foreach (var n in new[] { 1, 2 })
                {
                    mainPart
                        .AddNewPart<HeaderPart>($"rIdHeader{n}")
                        .PutXDocument(
                            XDocument.Parse(
                                $"<w:hdr {s_namespaces}><w:p><w:r><w:t>Header {n}</w:t></w:r></w:p></w:hdr>"
                            )
                        );
                    mainPart
                        .AddNewPart<FooterPart>($"rIdFooter{n}")
                        .PutXDocument(
                            XDocument.Parse(
                                $"<w:ftr {s_namespaces}><w:p><w:r><w:t>Footer {n}</w:t></w:r></w:p></w:ftr>"
                            )
                        );
                }
            }
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var sections = doc.MainDocumentPart!.GetXDocument().Descendants(W.sectPr).Select(DescribeSection);
        await Assert
            .That(string.Join(" | ", sections))
            .IsEqualTo(
                $"rIdHeader1 rIdFooter1 pgSz pgMar@{expectedTopMargin} | rIdHeader2 rIdFooter2 pgSz pgMar@{expectedTopMargin}"
            );
    }

    private static string TrackedSectionProperties(int n) =>
        $"""
            <w:sectPr>
              <w:headerReference w:type="default" r:id="rIdHeader{n}"/>
              <w:footerReference w:type="default" r:id="rIdFooter{n}"/>
              <w:pgSz w:w="11906" w:h="16838"/>
              <w:pgMar w:top="720" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
              <w:sectPrChange w:id="{n}" w:author="Author" w:date="2026-10-15T00:00:00Z">
                <w:sectPr>
                  <w:pgSz w:w="11906" w:h="16838"/>
                  <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
                </w:sectPr>
              </w:sectPrChange>
            </w:sectPr>
            """;

    private static string DescribeSection(XElement sectPr) =>
        string.Join(
            " ",
            sectPr
                .Elements()
                .Select(e =>
                    e.Name == W.pgMar
                        ? $"pgMar@{(string?)e.Attribute(W.top)}"
                        : (string?)e.Attribute(R.id) ?? e.Name.LocalName
                )
        );

    private static readonly string s_namespaces = $"xmlns:w=\"{W.w.NamespaceName}\" xmlns:r=\"{R.r.NamespaceName}\"";

    private static WmlDocument CreateDocument(string bodyXml, Action<MainDocumentPart>? addParts = null)
    {
        using var stream = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.PutXDocument(
                XDocument.Parse($"<w:document {s_namespaces}><w:body>{bodyXml}</w:body></w:document>")
            );
            addParts?.Invoke(mainPart);
        }
        return new WmlDocument("RevisionProcessorTest.docx", stream.ToArray());
    }
}
#pragma warning restore CA1707
