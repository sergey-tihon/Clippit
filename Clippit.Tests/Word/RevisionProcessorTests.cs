// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit.Word;

namespace Clippit.Tests.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
public class RevisionProcessorTests : TestsBase
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
    public async Task RP001(string name)
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

        await Assert.That(baselineAcceptedFi.Exists).IsTrue().WithMessage($"No Accepted baseline document for {name}");
        var acceptedResult = WmlComparer.Compare(
            new WmlDocument(baselineAcceptedFi.FullName),
            afterAcceptingWml,
            settings
        );
        var acceptedRevisions = WmlComparer.GetRevisions(acceptedResult, settings);
        await Assert
            .That(acceptedRevisions)
            .IsEmpty()
            .WithMessage($"Regression: Accepted output differs from baseline for {name}");

        await Assert.That(baselineRejectedFi.Exists).IsTrue().WithMessage($"No Rejected baseline document for {name}");
        var rejectedResult = WmlComparer.Compare(
            new WmlDocument(baselineRejectedFi.FullName),
            afterRejectingWml,
            settings
        );
        var rejectedRevisions = WmlComparer.GetRevisions(rejectedResult, settings);
        await Assert
            .That(rejectedRevisions)
            .IsEmpty()
            .WithMessage($"Regression: Rejected output differs from baseline for {name}");
    }
}
#pragma warning restore CA1707
