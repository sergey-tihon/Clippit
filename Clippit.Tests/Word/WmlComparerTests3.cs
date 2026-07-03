// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Drawing;
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word;

/// <summary>
/// Regression tests for defensive fixes in WmlComparer:
///   #380 — HashBlockLevelContent must not throw ArgumentNullException when a block-level element lacks pt:Unid.
///   #385 — FindIndexOfNextParaMark must not throw InvalidCastException when the ComparisonUnit[] contains
///           a non-ComparisonUnitWord element (e.g. ComparisonUnitGroup for a table).
/// </summary>
public class WmlComparerTests3 : TestsBase
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static void AddMinimalStylesPart(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();
        stylesPart.PutXDocument(
            new XDocument(
                new XElement(
                    W + "styles",
                    new XElement(
                        W + "style",
                        new XAttribute(W + "type", "paragraph"),
                        new XAttribute(W + "default", "1"),
                        new XAttribute(W + "styleId", "Normal"),
                        new XElement(W + "name", new XAttribute(W + "val", "Normal"))
                    )
                )
            )
        );
    }

    /// <summary>
    /// Creates a minimal WmlDocument containing the specified paragraphs.
    /// </summary>
    private static WmlDocument CreateDocxWithParagraphs(params string[] paragraphTexts)
    {
        using var ms = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            AddMinimalStylesPart(mainPart);
            mainPart.PutXDocument(
                new XDocument(
                    new XElement(
                        W + "document",
                        new XElement(
                            W + "body",
                            paragraphTexts
                                .Select(text => new XElement(
                                    W + "p",
                                    new XElement(W + "r", new XElement(W + "t", text))
                                ))
                                .Append(new XElement(W + "sectPr"))
                        )
                    )
                )
            );
        }

        return new WmlDocument("test.docx", ms.ToArray());
    }

    /// <summary>
    /// Creates a minimal WmlDocument containing a paragraph followed by a single-row table with the given cell texts.
    /// </summary>
    private static WmlDocument CreateDocxWithParagraphAndTable(string paragraphText, params string[] cellTexts)
    {
        using var ms = new MemoryStream();
        using (var wordDoc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            AddMinimalStylesPart(mainPart);
            mainPart.PutXDocument(
                new XDocument(
                    new XElement(
                        W + "document",
                        new XElement(
                            W + "body",
                            new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", paragraphText))),
                            new XElement(
                                W + "tbl",
                                new XElement(
                                    W + "tr",
                                    cellTexts.Select(cell => new XElement(
                                        W + "tc",
                                        new XElement(W + "p", new XElement(W + "r", new XElement(W + "t", cell)))
                                    ))
                                )
                            ),
                            new XElement(W + "sectPr")
                        )
                    )
                )
            );
        }

        return new WmlDocument("test.docx", ms.ToArray());
    }

    /// <summary>
    /// Regression test for issue #380.
    /// <para>
    /// <c>HashBlockLevelContent</c> used to call <c>.ToDictionary(d => (string)d.Attribute(PtOpenXml.Unid))</c>
    /// on block-level elements without first filtering out elements that lack the <c>pt:Unid</c> attribute.
    /// Any such element would inject a <c>null</c> key and cause an <c>ArgumentNullException</c>.
    /// </para>
    /// <para>
    /// This test exercises <c>Consolidate</c> with two revised documents (which calls <c>CompareInternal</c>
    /// with <c>preProcessMarkupInOriginal = false</c>, the code path where <c>HashBlockLevelContent</c> receives
    /// a document that may contain block-level elements without <c>pt:Unid</c> after revision acceptance).
    /// </para>
    /// </summary>
    [Test]
    public async Task WC380_Consolidate_MultipleRevisions_DoesNotThrow()
    {
        var original = CreateDocxWithParagraphs("Paragraph one", "Paragraph two", "Paragraph three");
        var revised1 = CreateDocxWithParagraphs("Paragraph one modified", "Paragraph two", "Paragraph three");
        var revised2 = CreateDocxWithParagraphs("Paragraph one", "Paragraph two modified", "Paragraph three");

        var settings = new WmlComparerSettings();
        var revisedDocList = new List<WmlRevisedDocumentInfo>
        {
            new()
            {
                RevisedDocument = revised1,
                Color = Color.LightBlue,
                Revisor = "Revisor1",
            },
            new()
            {
                RevisedDocument = revised2,
                Color = Color.LightGreen,
                Revisor = "Revisor2",
            },
        };

        // Must not throw ArgumentNullException from null pt:Unid keys in the block-level element dictionary.
        var result = WmlComparer.Consolidate(original, revisedDocList, settings);
        await Assert.That(result).IsNotNull();
    }

    /// <summary>
    /// Regression test for issue #385.
    /// <para>
    /// <c>FindIndexOfNextParaMark</c> used to hard-cast every element of the <c>ComparisonUnit[]</c> to
    /// <c>ComparisonUnitWord</c>.  When the array contains a <c>ComparisonUnitGroup</c> (produced for tables),
    /// this threw <c>InvalidCastException</c>.
    /// </para>
    /// <para>
    /// This test exercises <c>Compare</c> on documents where one has a paragraph followed by a table.
    /// The comparison algorithm's LCS pass produces a <c>ComparisonUnit[]</c> that mixes
    /// <c>ComparisonUnitWord</c> and <c>ComparisonUnitGroup</c> entries, which are passed to
    /// <c>FindIndexOfNextParaMark</c>.
    /// </para>
    /// </summary>
    [Test]
    public async Task WC385_Compare_TableNearParaMark_DoesNotThrow()
    {
        var source1 = CreateDocxWithParagraphAndTable("Original paragraph text", "Cell A1", "Cell A2");
        var source2 = CreateDocxWithParagraphAndTable("Modified paragraph text", "Cell A1 changed", "Cell A2");

        var settings = new WmlComparerSettings();

        // Must not throw InvalidCastException from ComparisonUnitGroup elements in FindIndexOfNextParaMark.
        var result = WmlComparer.Compare(source1, source2, settings);
        await Assert.That(result).IsNotNull();
    }
}
