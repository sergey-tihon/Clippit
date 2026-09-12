// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="AddDocxTextHelper.AppendParagraphToDocument"/>, which appends a
/// formatted paragraph (bold/italic/underline/colors/style) to the end of a Wordprocessing document body.
/// </summary>
public class AddDocxTextHelperTests : TestsBase
{
    private static readonly string TestFilesDir = "../../../../TestFiles/";
    private static readonly string DocxPath = Path.Combine(TestFilesDir, "Blank-wml.docx");

    private static Paragraph GetLastParagraph(WmlDocument doc)
    {
        using var streamDoc = new OpenXmlMemoryStreamDocument(doc);
        using var wDoc = streamDoc.GetWordprocessingDocument();
        return wDoc.MainDocumentPart.Document.Body.Elements<Paragraph>().Last();
    }

    [Test]
    public async Task ADT001_AppendParagraphToDocument_PlainText_AddsParagraphWithText()
    {
        var original = new WmlDocument(DocxPath);
        var result = AddDocxTextHelper.AppendParagraphToDocument(
            original,
            "Hello, world!",
            isBold: false,
            isItalic: false,
            isUnderline: false,
            foreColor: null,
            backColor: null,
            styleName: null
        );

        var paragraph = GetLastParagraph(result);
        var run = paragraph.Elements<Run>().Single();
        await Assert.That(run.Elements<Text>().Single().Text).IsEqualTo("Hello, world!");
        await Assert.That(run.RunProperties.HasChildren).IsFalse();
    }

    [Test]
    public async Task ADT002_AppendParagraphToDocument_Bold_SetsBoldRunProperty()
    {
        var original = new WmlDocument(DocxPath);
        var result = AddDocxTextHelper.AppendParagraphToDocument(
            original,
            "Bold text",
            isBold: true,
            isItalic: false,
            isUnderline: false,
            foreColor: null,
            backColor: null,
            styleName: null
        );

        var run = GetLastParagraph(result).Elements<Run>().Single();
        await Assert.That(run.RunProperties.Elements<Bold>().Any()).IsTrue();
        await Assert.That(run.RunProperties.Elements<Italic>().Any()).IsFalse();
    }

    [Test]
    public async Task ADT003_AppendParagraphToDocument_ItalicAndUnderline_SetsBothRunProperties()
    {
        var original = new WmlDocument(DocxPath);
        var result = AddDocxTextHelper.AppendParagraphToDocument(
            original,
            "Italic underline text",
            isBold: false,
            isItalic: true,
            isUnderline: true,
            foreColor: null,
            backColor: null,
            styleName: null
        );

        var run = GetLastParagraph(result).Elements<Run>().Single();
        await Assert.That(run.RunProperties.Elements<Italic>().Any()).IsTrue();
        var underline = run.RunProperties.Elements<Underline>().Single();
        await Assert.That(underline.Val.Value).IsEqualTo(UnderlineValues.Single);
    }

    [Test]
    public async Task ADT004_AppendParagraphToDocument_ForeColor_SetsColorRunProperty()
    {
        var original = new WmlDocument(DocxPath);
        var result = AddDocxTextHelper.AppendParagraphToDocument(
            original,
            "Colored text",
            isBold: false,
            isItalic: false,
            isUnderline: false,
            foreColor: "Red",
            backColor: null,
            styleName: null
        );

        var run = GetLastParagraph(result).Elements<Run>().Single();
        var color = run.RunProperties.Elements<Color>().Single();
        await Assert.That(color.Val.Value).IsEqualTo("ff0000");
    }

    [Test]
    public async Task ADT005_AppendParagraphToDocument_BackColor_SetsShadingRunProperty()
    {
        var original = new WmlDocument(DocxPath);
        var result = AddDocxTextHelper.AppendParagraphToDocument(
            original,
            "Shaded text",
            isBold: false,
            isItalic: false,
            isUnderline: false,
            foreColor: null,
            backColor: "Yellow",
            styleName: null
        );

        var run = GetLastParagraph(result).Elements<Run>().Single();
        var shading = run.RunProperties.Elements<Shading>().Single();
        await Assert.That(shading.Val.Value).IsEqualTo(ShadingPatternValues.Clear);
        await Assert.That(shading.Fill.Value).IsEqualTo("ffff00");
    }

    [Test]
    [Arguments("NotARealColor", null)]
    [Arguments(null, "NotARealColor")]
    public async Task ADT006_AppendParagraphToDocument_InvalidColor_Throws(string? foreColor, string? backColor)
    {
        var original = new WmlDocument(DocxPath);
        await Assert
            .That(() =>
                AddDocxTextHelper.AppendParagraphToDocument(
                    original,
                    "Bad color",
                    isBold: false,
                    isItalic: false,
                    isUnderline: false,
                    foreColor: foreColor,
                    backColor: backColor,
                    styleName: null
                )
            )
            .Throws<OpenXmlPowerToolsException>();
    }

    [Test]
    public async Task ADT007_AppendParagraphToDocument_ExistingStyleName_AppliesParagraphStyleId()
    {
        var original = new WmlDocument(DocxPath);
        var result = AddDocxTextHelper.AppendParagraphToDocument(
            original,
            "Styled text",
            isBold: false,
            isItalic: false,
            isUnderline: false,
            foreColor: null,
            backColor: null,
            styleName: "Heading1"
        );

        var paragraph = GetLastParagraph(result);
        await Assert.That(paragraph.ParagraphProperties.ParagraphStyleId.Val.Value).IsEqualTo("Heading1");
    }

    [Test]
    public async Task ADT008_AppendParagraphToDocument_UnknownStyleName_Throws()
    {
        var original = new WmlDocument(DocxPath);
        await Assert
            .That(() =>
                AddDocxTextHelper.AppendParagraphToDocument(
                    original,
                    "Bad style",
                    isBold: false,
                    isItalic: false,
                    isUnderline: false,
                    foreColor: null,
                    backColor: null,
                    styleName: "ThisStyleDoesNotExist"
                )
            )
            .Throws<OpenXmlPowerToolsException>();
    }

    [Test]
    public async Task ADT009_AppendParagraphToDocument_ProducesValidDocument()
    {
        var original = new WmlDocument(DocxPath);
        var result = AddDocxTextHelper.AppendParagraphToDocument(
            original,
            "Validated paragraph",
            isBold: true,
            isItalic: true,
            isUnderline: true,
            foreColor: "Blue",
            backColor: "Green",
            styleName: null
        );

        using var streamDoc = new OpenXmlMemoryStreamDocument(result);
        using var wDoc = streamDoc.GetWordprocessingDocument();
        await Validate(wDoc);
    }
}
