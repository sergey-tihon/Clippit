// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
public partial class RevisionProcessorTests
{
    [Test]
    [Arguments("acceptParagraph", true, "both")]
    [Arguments("acceptParagraph", false, "ids")]
    [Arguments("rejectParagraph", true, "both")]
    [Arguments("rejectRow", true, "parent")]
    [Arguments("rejectCell", true, "ids")]
    public async Task RP066_CommentMetadataKeepsRemovedTerminalParagraphIdentity(
        string scenario,
        bool keepParagraph,
        string metadata
    )
    {
        const string revisionAttributes = "w:id=\"1\" w:author=\"Author\" w:date=\"2026-10-15T00:00:00Z\"";
        const string keptParagraph = """
            <w:p w14:paraId="11111111"><w:r><w:t>Parent kept</w:t></w:r></w:p>
            """;
        var reject = scenario.StartsWith("reject", StringComparison.Ordinal);
        var revision = reject ? "ins" : "del";
        var textElement = reject ? "t" : "delText";
        var finalParagraph = $"""
            <w:p w14:paraId="22222222">
              <w:pPr><w:rPr><w:{revision} {revisionAttributes}/></w:rPr></w:pPr>
              <w:{revision} w:id="2" w:author="Author" w:date="2026-10-15T00:00:00Z"><w:r><w:{textElement}>Parent removed</w:{textElement}></w:r></w:{revision}>
            </w:p>
            """;
        const string structuralFinalParagraph = """
            <w:p w14:paraId="22222222"><w:r><w:t>Parent removed</w:t></w:r></w:p>
            """;
        var parentContent = scenario switch
        {
            "acceptParagraph" or "rejectParagraph" => (keepParagraph ? keptParagraph : "") + finalParagraph,
            "rejectRow" => $"""
                <w:tbl>
                  <w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>
                  <w:tr><w:tc>{keptParagraph}</w:tc></w:tr>
                  <w:tr><w:trPr><w:ins {revisionAttributes}/></w:trPr><w:tc>{structuralFinalParagraph}</w:tc></w:tr>
                </w:tbl>
                """,
            "rejectCell" => $"""
                <w:tbl>
                  <w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>
                  <w:tr>
                    <w:tc>{keptParagraph}</w:tc>
                    <w:tc><w:tcPr><w:cellIns {revisionAttributes}/></w:tcPr>{structuralFinalParagraph}</w:tc>
                  </w:tr>
                </w:tbl>
                """,
            _ => throw new ArgumentOutOfRangeException(nameof(scenario)),
        };
        XNamespace cid = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
        var parentExtension = metadata == "both" ? """<w15:commentEx w15:paraId="22222222" w15:done="1"/>""" : "";
        var parentDurableId =
            metadata != "parent" ? """<w16cid:commentId w16cid:paraId="22222222" w16cid:durableId="00000001"/>""" : "";
        // The parent-only case is referenced exclusively through the reply's paraIdParent;
        // the ids-only case has no extended-comment reference to the revised comment.
        var replyParent = metadata != "ids" ? " w15:paraIdParent=\"22222222\"" : "";
        var extendedXml = XDocument.Parse(
            $"""
            <w15:commentsEx xmlns:w15="{W15.w15}">
              {parentExtension}
              <w15:commentEx w15:paraId="33333333"{replyParent} w15:done="0"/>
            </w15:commentsEx>
            """
        );
        var idsXml = XDocument.Parse(
            $"""
            <w16cid:commentsIds xmlns:w16cid="{cid}">
              {parentDurableId}
              <w16cid:commentId w16cid:paraId="33333333" w16cid:durableId="00000002"/>
            </w16cid:commentsIds>
            """
        );
        var source = CreateDocument(
            """
            <w:p>
              <w:commentRangeStart w:id="0"/><w:r><w:t>Commented</w:t></w:r>
              <w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r>
              <w:commentRangeStart w:id="2"/><w:r><w:t>Unchanged</w:t></w:r>
              <w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r>
            </w:p>
            """,
            mainPart =>
            {
                mainPart
                    .AddNewPart<WordprocessingCommentsPart>()
                    .PutXDocument(
                        XDocument.Parse(
                            $"""
                            <w:comments xmlns:w="{W.w}" xmlns:w14="{W14.w14}"
                                xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14">
                              <w:comment w:id="0" w:author="Author" w:date="2026-10-15T00:00:00Z">{parentContent}</w:comment>
                              <w:comment w:id="1" w:author="Reply author" w:date="2026-10-15T00:00:00Z">
                                <w:p w14:paraId="33333333"><w:r><w:t>Reply unchanged</w:t></w:r></w:p>
                              </w:comment>
                              <w:comment w:id="2" w:author="Author" w:date="2026-10-15T00:00:00Z">
                                <w:p w14:paraId="44444444"><w:r><w:t>Unreferenced first</w:t></w:r></w:p>
                                <w:p w14:paraId="55555555"><w:r><w:t>Unreferenced last</w:t></w:r></w:p>
                              </w:comment>
                            </w:comments>
                            """
                        )
                    );
                mainPart.AddNewPart<WordprocessingCommentsExPart>().PutXDocument(new XDocument(extendedXml));
                mainPart.AddNewPart<WordprocessingCommentsIdsPart>().PutXDocument(new XDocument(idsXml));
            }
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc, FileFormatVersions.Office2019);
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var mainPart = doc.MainDocumentPart!;
        var comments = mainPart.WordprocessingCommentsPart!.GetXDocument().Root!.Elements(W.comment).ToList();
        var parent = comments.Single(c => (string?)c.Attribute(W.id) == "0");
        var paragraphs = parent.Descendants(W.p).ToList();
        await Assert.That(paragraphs.Count).IsEqualTo(1);
        await Assert.That(TextOf(paragraphs[0])).IsEqualTo(keepParagraph ? "Parent kept" : "");
        await Assert.That((string?)paragraphs[0].Attribute(W14.paraId)).IsEqualTo("22222222");
        var reply = comments.Single(c => (string?)c.Attribute(W.id) == "1");
        await Assert.That(TextOf(reply)).IsEqualTo("Reply unchanged");
        await Assert.That((string?)reply.Descendants(W.p).Single().Attribute(W14.paraId)).IsEqualTo("33333333");
        var unreferenced = comments.Single(c => (string?)c.Attribute(W.id) == "2");
        await Assert
            .That(string.Join("|", unreferenced.Descendants(W.p).Select(TextOf)))
            .IsEqualTo("Unreferenced first|Unreferenced last");
        await Assert
            .That(string.Join("|", unreferenced.Descendants(W.p).Select(p => (string?)p.Attribute(W14.paraId))))
            .IsEqualTo("44444444|55555555");

        var extended = mainPart.WordprocessingCommentsExPart!.GetXDocument();
        var ids = mainPart.WordprocessingCommentsIdsPart!.GetXDocument();
        await Assert
            .That(extended.ToString(SaveOptions.DisableFormatting))
            .IsEqualTo(extendedXml.ToString(SaveOptions.DisableFormatting));
        await Assert
            .That(ids.ToString(SaveOptions.DisableFormatting))
            .IsEqualTo(idsXml.ToString(SaveOptions.DisableFormatting));
        var terminalIds = comments.Select(c => (string?)c.Descendants(W.p).Last().Attribute(W14.paraId)).ToHashSet();
        foreach (
            var reference in extended
                .Descendants()
                .Attributes()
                .Where(a => a.Name == W15.w15 + "paraId" || a.Name == W15.w15 + "paraIdParent")
        )
        {
            await Assert.That(terminalIds.Contains(reference.Value)).IsTrue();
        }
        foreach (var reference in ids.Descendants().Attributes(cid + "paraId"))
        {
            await Assert.That(terminalIds.Contains(reference.Value)).IsTrue();
        }
    }
}
#pragma warning restore CA1707
