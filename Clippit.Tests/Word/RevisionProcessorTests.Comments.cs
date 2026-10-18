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
    [Arguments("header", false, "Alpha delta gamma")]
    [Arguments("header", true, "Alpha beta| gamma")]
    [Arguments("footer", false, "Alpha delta gamma")]
    [Arguments("footer", true, "Alpha beta| gamma")]
    [Arguments("footnote", false, "Alpha delta gamma")]
    [Arguments("footnote", true, "Alpha beta| gamma")]
    [Arguments("endnote", false, "Alpha delta gamma")]
    [Arguments("endnote", true, "Alpha beta| gamma")]
    [Arguments("comment", false, "Alpha delta gamma")]
    [Arguments("comment", true, "Alpha beta| gamma")]
    public async Task RP057_RevisionsInEveryStory(string story, bool reject, string expectedParagraphs)
    {
        const string storyId = "rIdStory";
        const string revisedParagraphs = """
            <w:p>
              <w:pPr><w:rPr><w:del w:id="11" w:author="Author" w:date="2026-10-15T00:00:00Z"/></w:rPr></w:pPr>
              <w:r><w:t xml:space="preserve">Alpha </w:t></w:r>
              <w:del w:id="12" w:author="Author" w:date="2026-10-15T00:00:00Z">
                <w:r><w:delText>beta</w:delText></w:r>
              </w:del>
              <w:ins w:id="13" w:author="Author" w:date="2026-10-15T00:00:00Z">
                <w:r><w:t>delta</w:t></w:r>
              </w:ins>
            </w:p>
            <w:p><w:r><w:t xml:space="preserve"> gamma</w:t></w:r></w:p>
            """;
        var (bodyXml, storyXml) = story switch
        {
            "header" => (
                $"""<w:p/><w:sectPr><w:headerReference w:type="default" r:id="{storyId}"/></w:sectPr>""",
                $"<w:hdr {s_namespaces}>{revisedParagraphs}</w:hdr>"
            ),
            "footer" => (
                $"""<w:p/><w:sectPr><w:footerReference w:type="default" r:id="{storyId}"/></w:sectPr>""",
                $"<w:ftr {s_namespaces}>{revisedParagraphs}</w:ftr>"
            ),
            "footnote" => (
                """<w:p><w:r><w:footnoteReference w:id="1"/></w:r></w:p>""",
                $"""<w:footnotes {s_namespaces}><w:footnote w:id="1">{revisedParagraphs}</w:footnote></w:footnotes>"""
            ),
            "endnote" => (
                """<w:p><w:r><w:endnoteReference w:id="1"/></w:r></w:p>""",
                $"""<w:endnotes {s_namespaces}><w:endnote w:id="1">{revisedParagraphs}</w:endnote></w:endnotes>"""
            ),
            "comment" => (
                """
                <w:p>
                  <w:commentRangeStart w:id="0"/>
                  <w:r><w:t>Commented</w:t></w:r>
                  <w:commentRangeEnd w:id="0"/>
                  <w:r><w:commentReference w:id="0"/></w:r>
                </w:p>
                """,
                $"""
                <w:comments {s_namespaces}>
                  <w:comment w:id="0" w:author="Author" w:date="2026-10-15T00:00:00Z" w:initials="A">
                    {revisedParagraphs}
                  </w:comment>
                </w:comments>
                """
            ),
            _ => throw new ArgumentOutOfRangeException(nameof(story)),
        };
        var source = CreateDocument(
            bodyXml,
            mainPart =>
            {
                OpenXmlPart storyPart = story switch
                {
                    "header" => mainPart.AddNewPart<HeaderPart>(storyId),
                    "footer" => mainPart.AddNewPart<FooterPart>(storyId),
                    "footnote" => mainPart.AddNewPart<FootnotesPart>(storyId),
                    "endnote" => mainPart.AddNewPart<EndnotesPart>(storyId),
                    _ => mainPart.AddNewPart<WordprocessingCommentsPart>(storyId),
                };
                storyPart.PutXDocument(XDocument.Parse(storyXml));
            }
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        var processedStory = doc.MainDocumentPart!.GetPartById(storyId);
        var paragraphs = processedStory.GetXDocument().Descendants(W.p).Select(TextOf);
        await Assert.That(string.Join("|", paragraphs)).IsEqualTo(expectedParagraphs);
        await Assert.That(RevisionProcessor.PartHasTrackedRevisions(processedStory)).IsFalse();
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
    }

    [Test]
    [Arguments("merge", false, "Parent first paragraph Parent final paragraph")]
    [Arguments("merge", true, "Parent first paragraph Parent final paragraph")]
    public async Task RP063_CommentKeepsLastParagraphId(string scenario, bool reject, string expectedParagraphs)
    {
        var commentParagraphs = scenario switch
        {
            "merge" =>
                $"""
                <w:p w14:paraId="11111111">
                  <w:pPr><w:rPr><w:{(
                    reject ? "ins" : "del"
                )} w:id="1" w:author="Author" w:date="2026-10-15T00:00:00Z"/></w:rPr></w:pPr>
                  <w:r><w:t>Parent first paragraph</w:t></w:r>
                </w:p>
                <w:p w14:paraId="22222222"><w:r><w:t xml:space="preserve"> Parent final paragraph</w:t></w:r></w:p>
                """,
            _ => throw new ArgumentOutOfRangeException(nameof(scenario)),
        };
        var commentNamespaces = $"xmlns:w=\"{W.w.NamespaceName}\" xmlns:w14=\"{W14.w14.NamespaceName}\"";
        var source = CreateDocument(
            """
            <w:p>
              <w:commentRangeStart w:id="0"/>
              <w:r><w:t>Commented</w:t></w:r>
              <w:commentRangeEnd w:id="0"/>
              <w:r><w:commentReference w:id="0"/></w:r>
            </w:p>
            """,
            mainPart =>
            {
                var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>("rIdComments");
                commentsPart.PutXDocument(
                    XDocument.Parse(
                        $"""<w:comments {commentNamespaces}><w:comment w:id="0" w:author="Author" w:date="2026-10-15T00:00:00Z">{commentParagraphs}</w:comment></w:comments>"""
                    )
                );
            }
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc, [$"The '{W14.w14.NamespaceName}:paraId' attribute is not declared."]);
        await Assert.That(RevisionProcessor.HasTrackedRevisions(doc)).IsFalse();
        var comments = doc.MainDocumentPart!.GetPartById("rIdComments").GetXDocument();
        var comment = comments.Root!.Element(W.comment)!;
        await Assert.That(string.Join("|", comment.Elements(W.p).Select(TextOf))).IsEqualTo(expectedParagraphs);
        await Assert.That((string?)comment.Elements(W.p).Last().Attribute(W14.paraId)).IsEqualTo("22222222");
    }

    [Test]
    [Arguments("body", "unchanged", false)]
    [Arguments("body", "unchanged", true)]
    [Arguments("body", "mergeElsewhere", false)]
    [Arguments("body", "mergeElsewhere", true)]
    [Arguments("body", "markerInsideMerge", false)]
    [Arguments("body", "markerInsideMerge", true)]
    [Arguments("body", "mergeIntoCustomXml", false)]
    [Arguments("body", "mergeIntoCustomXml", true)]
    [Arguments("comment", "unchanged", false)]
    [Arguments("comment", "unchanged", true)]
    [Arguments("comment", "mergeElsewhere", false)]
    [Arguments("comment", "mergeElsewhere", true)]
    [Arguments("comment", "markerInsideMerge", false)]
    [Arguments("comment", "markerInsideMerge", true)]
    [Arguments("comment", "mergeIntoCustomXml", false)]
    [Arguments("comment", "mergeIntoCustomXml", true)]
    public async Task RP064_RebuiltContainerKeepsBlockMarkup(string story, string scenario, bool reject)
    {
        const string bookmarkStart = "<w:bookmarkStart w:id=\"1\" w:name=\"Kept\"/>";
        const string bookmarkEnd = "<w:bookmarkEnd w:id=\"1\"/>";
        const string unchanged = """
            <w:bookmarkStart w:id="1" w:name="Kept"/><w:p><w:r><w:t>Bookmarked</w:t></w:r></w:p><w:bookmarkEnd w:id="1"/>
            <w:customXml w:element="payload"><w:p><w:r><w:t>Wrapped</w:t></w:r></w:p></w:customXml>
            <w:sdt><w:sdtPr><w:tag w:val="empty"/></w:sdtPr><w:sdtContent><w:p/></w:sdtContent></w:sdt>
            """;
        string P(string text) => $"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>";
        string Pm(string text) =>
            $"<w:p><w:pPr><w:rPr><w:{(reject ? "ins" : "del")} w:id=\"9\" w:author=\"Author\" w:date=\"2026-10-15T00:00:00Z\"/></w:rPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>";

        var blockXml = scenario switch
        {
            "unchanged" => unchanged,
            "mergeElsewhere" =>
                $"{bookmarkStart}{P("A")}{bookmarkEnd}<w:customXml w:element=\"payload\">{P("W")}</w:customXml>{Pm("B")}{P("C")}",
            "markerInsideMerge" => $"{Pm("B")}{bookmarkStart}{P("C")}{bookmarkEnd}",
            "mergeIntoCustomXml" =>
                $"{Pm("First")}<w:customXml w:element=\"c\">{P("Second")}</w:customXml>{P("Third")}",
            _ => throw new ArgumentOutOfRangeException(nameof(scenario)),
        };
        var expectedXml = scenario switch
        {
            "unchanged" => unchanged,
            "mergeElsewhere" =>
                $"{bookmarkStart}{P("A")}{bookmarkEnd}<w:customXml w:element=\"payload\">{P("W")}</w:customXml><w:p><w:r><w:t>B</w:t></w:r><w:r><w:t>C</w:t></w:r></w:p>",
            "markerInsideMerge" =>
                $"<w:p><w:r><w:t>B</w:t></w:r>{bookmarkStart}<w:r><w:t>C</w:t></w:r></w:p>{bookmarkEnd}",
            "mergeIntoCustomXml" => $"<w:p><w:r><w:t>First</w:t></w:r><w:r><w:t>Second</w:t></w:r></w:p>{P("Third")}",
            _ => throw new ArgumentOutOfRangeException(nameof(scenario)),
        };
        var bodyXml = story switch
        {
            "body" => blockXml,
            "comment" => """
                <w:p>
                  <w:commentRangeStart w:id="0"/>
                  <w:r><w:t>Commented</w:t></w:r>
                  <w:commentRangeEnd w:id="0"/>
                  <w:r><w:commentReference w:id="0"/></w:r>
                </w:p>
                """,
            _ => throw new ArgumentOutOfRangeException(nameof(story)),
        };
        var source = CreateDocument(
            bodyXml,
            mainPart =>
            {
                if (story != "comment")
                    return;
                var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>("rIdComments");
                commentsPart.PutXDocument(
                    XDocument.Parse(
                        $"<w:comments {s_namespaces}><w:comment w:id=\"0\" w:author=\"Author\">{blockXml}</w:comment></w:comments>"
                    )
                );
            }
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        var processedContainer =
            story == "body"
                ? doc.MainDocumentPart!.GetXDocument().Root!.Element(W.body)!
                : doc.MainDocumentPart!.GetPartById("rIdComments").GetXDocument().Root!.Element(W.comment)!;
        var expectedContainer = XElement.Parse($"<w:x {s_namespaces}>{expectedXml}</w:x>");
        await Assert
            .That(string.Concat(processedContainer.Elements().Select(e => e.ToString(SaveOptions.DisableFormatting))))
            .IsEqualTo(
                string.Concat(expectedContainer.Elements().Select(e => e.ToString(SaveOptions.DisableFormatting)))
            );
    }

    [Test]
    [Arguments(false)]
    [Arguments(true)]
    public async Task RP065_AltChunkOrderIsPreservedWhenMergingCommentParagraphs(bool reject)
    {
        var mark = reject ? "ins" : "del";
        var commentXml =
            $"<w:p><w:pPr><w:rPr><w:{mark} w:id=\"1\" w:author=\"Author\" w:date=\"2026-10-15T00:00:00Z\"/></w:rPr></w:pPr><w:r><w:t>Before</w:t></w:r></w:p>"
            + "<w:altChunk r:id=\"rIdChunk\"/>"
            + "<w:p><w:r><w:t>After</w:t></w:r></w:p>";
        var source = CreateDocument(
            "<w:p/>",
            mainPart =>
            {
                var commentsPart = mainPart.AddNewPart<WordprocessingCommentsPart>("rIdComments");
                var altChunkPart = commentsPart.AddAlternativeFormatImportPart(
                    AlternativeFormatImportPartType.TextPlain,
                    "rIdChunk"
                );
                using var altChunkStream = altChunkPart.GetStream(FileMode.Create);
                using var altChunkWriter = new StreamWriter(altChunkStream);
                altChunkWriter.Write("chunk");
                commentsPart.PutXDocument(
                    XDocument.Parse(
                        $"<w:comments {s_namespaces}><w:comment w:id=\"0\" w:author=\"Author\">{commentXml}</w:comment></w:comments>"
                    )
                );
            }
        );

        var processed = reject ? RevisionProcessor.RejectRevisions(source) : RevisionProcessor.AcceptRevisions(source);

        using var streamDoc = new OpenXmlMemoryStreamDocument(processed);
        using var doc = streamDoc.GetWordprocessingDocument();
        await Validate(doc);
        var comment = doc.MainDocumentPart!.WordprocessingCommentsPart!.GetXDocument().Root!.Element(W.comment)!;
        await Assert.That(string.Join("|", comment.Elements().Select(e => e.Name.LocalName))).IsEqualTo("p|altChunk|p");
        await Assert.That(string.Join("|", comment.Elements(W.p).Select(TextOf))).IsEqualTo("Before|After");
    }
}
#pragma warning restore CA1707
