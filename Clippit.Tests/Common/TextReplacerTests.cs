// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Internal;
using Clippit.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

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
}
