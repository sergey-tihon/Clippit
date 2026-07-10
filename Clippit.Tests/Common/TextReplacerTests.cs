// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit;
using Clippit.Internal;
using Clippit.PowerPoint;
using Clippit.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="TextReplacer.SearchAndReplace(WmlDocument, string, string, bool)"/>.
/// These tests verify that the replacement actually modifies document text and that
/// case-sensitivity is honoured — the existing sample tests lack assertions entirely.
/// </summary>
public class TextReplacerTests
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>Creates a minimal WmlDocument containing a single paragraph with the given text.</summary>
    private static WmlDocument CreateDocxWithText(string text)
    {
        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                // TextReplacer checks DocumentSettingsPart for trackRevisions — the part must exist.
                var settingsPart = mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.PutXDocument(new XDocument(new XElement(W + "settings")));
                mainPart.PutXDocument(
                    new XDocument(
                        new XElement(
                            W + "document",
                            new XElement(
                                W + "body",
                                new XElement(
                                    W + "p",
                                    new XElement(
                                        W + "r",
                                        new XElement(
                                            W + "t",
                                            new XAttribute(XNamespace.Xml + "space", "preserve"),
                                            text
                                        )
                                    )
                                )
                            )
                        )
                    )
                );
            }
            bytes = ms.ToArray();
        }
        return new WmlDocument("test.docx", bytes);
    }

    /// <summary>Reads back all w:t text from the main document part and concatenates it.</summary>
    private static string GetDocumentText(WmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var wordDoc = WordprocessingDocument.Open(ms, false);
        return string.Concat(wordDoc.MainDocumentPart!.GetXDocument().Descendants(W + "t").Select(t => (string)t));
    }

    [Test]
    public async Task TR001_SearchAndReplace_MatchFound_ReplacesText()
    {
        var doc = CreateDocxWithText("Hello World");
        var result = TextReplacer.SearchAndReplace(doc, "World", "OpenXml", true);
        await Assert.That(GetDocumentText(result)).IsEqualTo("Hello OpenXml");
    }

    [Test]
    public async Task TR002_SearchAndReplace_NoMatch_ReturnsOriginalText()
    {
        var doc = CreateDocxWithText("Hello World");
        var result = TextReplacer.SearchAndReplace(doc, "NotPresent", "Replacement", true);
        await Assert.That(GetDocumentText(result)).IsEqualTo("Hello World");
    }

    [Test]
    public async Task TR003_SearchAndReplace_CaseInsensitive_MatchesAnyCase()
    {
        var doc = CreateDocxWithText("Hello World");
        var result = TextReplacer.SearchAndReplace(doc, "hello", "Hi", false);
        await Assert.That(GetDocumentText(result)).IsEqualTo("Hi World");
    }

    [Test]
    public async Task TR004_SearchAndReplace_CaseSensitive_DoesNotMatchWrongCase()
    {
        var doc = CreateDocxWithText("Hello World");
        var result = TextReplacer.SearchAndReplace(doc, "hello", "Hi", true);
        // matchCase=true: "hello" should NOT match "Hello"
        await Assert.That(GetDocumentText(result)).IsEqualTo("Hello World");
    }

    [Test]
    public async Task TR005_SearchAndReplace_OriginalDocumentIsNotMutated()
    {
        var doc = CreateDocxWithText("Hello World");
        _ = TextReplacer.SearchAndReplace(doc, "World", "OpenXml", true);
        // SearchAndReplace returns a new document — the original must be unchanged
        await Assert.That(GetDocumentText(doc)).IsEqualTo("Hello World");
    }

    [Test]
    public async Task TR006_SearchAndReplace_ReplaceEntireText()
    {
        var doc = CreateDocxWithText("Find me");
        var result = TextReplacer.SearchAndReplace(doc, "Find me", "Found it", true);
        await Assert.That(GetDocumentText(result)).IsEqualTo("Found it");
    }

    /// <summary>
    /// Regression test for issue #381 — replacing a match with an empty string must not
    /// throw <see cref="IndexOutOfRangeException"/> when consolidating adjacent runs.
    /// </summary>
    [Test]
    public async Task TR007_SearchAndReplace_EmptyReplacement_DoesNotThrow()
    {
        var doc = CreateDocxWithText("Remove this word");
        // Replacing with "" previously crashed with IndexOutOfRangeException
        var result = TextReplacer.SearchAndReplace(doc, "Remove this word", "", true);
        var text = GetDocumentText(result);
        await Assert.That(text).IsEqualTo("");
    }

    /// <summary>
    /// Variant of TR007: replace a partial match (not the whole run) with empty string.
    /// </summary>
    [Test]
    public async Task TR008_SearchAndReplace_PartialMatchEmptyReplacement_DoesNotThrow()
    {
        var doc = CreateDocxWithText("Hello World");
        var result = TextReplacer.SearchAndReplace(doc, "Hello ", "", true);
        var text = GetDocumentText(result);
        await Assert.That(text).IsEqualTo("World");
    }

    // ── PPTX overload tests ───────────────────────────────────────────────────

    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    /// <summary>
    /// Creates a minimal in-memory PmlDocument with a single slide containing the given text.
    /// </summary>
    private static PmlDocument CreatePptxWithText(string text)
    {
        byte[] bytes;
        using (var ms = new MemoryStream())
        {
            using (var pptx = PresentationDocument.Create(ms, PresentationDocumentType.Presentation))
            {
                var pressPart = pptx.AddPresentationPart();

                // Minimal Presentation element
                pressPart.Presentation = new Presentation(
                    new SlideIdList(),
                    new SlideSize { Cx = 9144000, Cy = 5143500 },
                    new NotesSize { Cx = 6858000, Cy = 9144000 }
                );

                var slidePart = pressPart.AddNewPart<SlidePart>();
                var slideIdList = pressPart.Presentation.SlideIdList
                    ?? throw new InvalidOperationException("Presentation.SlideIdList was not initialized.");
                var slideId = slideIdList.AppendChild(new SlideId());
                slideId.Id = 256;
                slideId.RelationshipId = pressPart.GetIdOfPart(slidePart);

                // Minimal slide XML with a text run containing 'text'
                var slideXml = new XDocument(
                    new XElement(
                        P + "sld",
                        new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
                        new XAttribute(XNamespace.Xmlns + "p", P.NamespaceName),
                        new XElement(
                            P + "cSld",
                            new XElement(
                                P + "spTree",
                                new XElement(
                                    P + "sp",
                                    new XElement(
                                        P + "txBody",
                                        new XElement(A + "p", new XElement(A + "r", new XElement(A + "t", text)))
                                    )
                                )
                            )
                        )
                    )
                );
                slidePart.PutXDocument(slideXml);
                pressPart.Presentation.Save();
            }
            bytes = ms.ToArray();
        }
        return new PmlDocument("test.pptx", bytes);
    }

    /// <summary>Reads back all a:t text from all slides and concatenates it.</summary>
    private static string GetPptxText(PmlDocument doc)
    {
        using var ms = new MemoryStream(doc.DocumentByteArray);
        using var pptx = PresentationDocument.Open(ms, false);
        return string.Concat(
            pptx.PresentationPart!.SlideParts.SelectMany(sp => sp.GetXDocument().Descendants(A + "t"))
                .Select(t => (string)t)
        );
    }

    [Test]
    public async Task TR009_Pptx_SearchAndReplace_MatchFound_ReplacesText()
    {
        var doc = CreatePptxWithText("Hello Presentation");
        var result = TextReplacer.SearchAndReplace(doc, "Presentation", "World", true);
        await Assert.That(GetPptxText(result)).IsEqualTo("Hello World");
    }

    [Test]
    public async Task TR010_Pptx_SearchAndReplace_CaseInsensitive_MatchesAnyCase()
    {
        var doc = CreatePptxWithText("Hello Presentation");
        var result = TextReplacer.SearchAndReplace(doc, "hello", "Hi", false);
        await Assert.That(GetPptxText(result)).IsEqualTo("Hi Presentation");
    }

    [Test]
    public async Task TR011_Pptx_SearchAndReplace_CaseSensitive_DoesNotMatchWrongCase()
    {
        var doc = CreatePptxWithText("Hello Presentation");
        var result = TextReplacer.SearchAndReplace(doc, "hello", "Hi", true);
        await Assert.That(GetPptxText(result)).IsEqualTo("Hello Presentation");
    }

    [Test]
    public async Task TR012_Pptx_SearchAndReplace_NoMatch_ReturnsOriginalText()
    {
        var doc = CreatePptxWithText("Hello Presentation");
        var result = TextReplacer.SearchAndReplace(doc, "NotPresent", "X", true);
        await Assert.That(GetPptxText(result)).IsEqualTo("Hello Presentation");
    }
}
