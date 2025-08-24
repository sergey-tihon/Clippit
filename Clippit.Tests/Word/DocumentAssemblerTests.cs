// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word;

public class DocumentAssemblerTests : TestsBase
{
    private readonly DirectoryInfo _sourceDir = new("../../../../TestFiles/DA/");

    [Test]
    [Arguments("DA001-TemplateDocument.docx", "DA-Data.xml", false)]
    [Arguments("DA002-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
    [Arguments("DA003-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
    [Arguments("DA004-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
    [Arguments("DA005-SelectRowData-NoData.docx", "DA-Data.xml", true)]
    [Arguments("DA006-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
    [Arguments("DA007-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
    [Arguments("DA008-TableElementWithNoTable.docx", "DA-Data.xml", true)]
    [Arguments("DA009-InvalidXPath.docx", "DA-Data.xml", true)]
    [Arguments("DA010-InvalidXml.docx", "DA-Data.xml", true)]
    [Arguments("DA011-SchemaError.docx", "DA-Data.xml", true)]
    [Arguments("DA012-OtherMarkupTypes.docx", "DA-Data.xml", true)]
    [Arguments("DA013-Runs.docx", "DA-Data.xml", false)]
    [Arguments("DA014-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
    [Arguments("DA015-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
    [Arguments("DA016-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
    [Arguments("DA017-FiveRuns.docx", "DA-Data.xml", true)]
    [Arguments("DA018-SmartQuotes.docx", "DA-Data.xml", false)]
    [Arguments("DA019-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
    [Arguments("DA020-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
    [Arguments("DA021-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
    [Arguments("DA022-InvalidXPath.docx", "DA-Data.xml", true)]
    [Arguments("DA023-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
    [Arguments("DA026-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
    [Arguments("DA027-XPathErrorInPara.docx", "DA-Data.xml", true)]
    [Arguments("DA028-NoPrototypeRow.docx", "DA-Data.xml", true)]
    [Arguments("DA029-NoDataForCell.docx", "DA-Data.xml", true)]
    [Arguments("DA030-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", false)] // Clippit support multi-value XPath in table cells
    [Arguments("DA031-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
    [Arguments("DA032-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
    [Arguments("DA033-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
    [Arguments("DA034-HeaderFooter.docx", "DA-Data.xml", false)]
    [Arguments("DA035-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
    [Arguments("DA036-SchemaErrorInConditional.docx", "DA-Data.xml", true)]
    [Arguments("DA100-TemplateDocument.docx", "DA-Data.xml", false)]
    [Arguments("DA101-TemplateDocument.docx", "DA-Data.xml", true)]
    [Arguments("DA102-TemplateDocument.docx", "DA-Data.xml", true)]
    [Arguments("DA201-TemplateDocument.docx", "DA-Data.xml", false)]
    [Arguments("DA202-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
    [Arguments("DA203-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
    [Arguments("DA204-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
    [Arguments("DA205-SelectRowData-NoData.docx", "DA-Data.xml", true)]
    [Arguments("DA206-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
    [Arguments("DA207-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
    [Arguments("DA209-InvalidXPath.docx", "DA-Data.xml", true)]
    [Arguments("DA210-InvalidXml.docx", "DA-Data.xml", true)]
    [Arguments("DA211-SchemaError.docx", "DA-Data.xml", true)]
    [Arguments("DA212-OtherMarkupTypes.docx", "DA-Data.xml", true)]
    [Arguments("DA213-Runs.docx", "DA-Data.xml", false)]
    [Arguments("DA214-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
    [Arguments("DA215-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
    [Arguments("DA216-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
    [Arguments("DA217-FiveRuns.docx", "DA-Data.xml", true)]
    [Arguments("DA218-SmartQuotes.docx", "DA-Data.xml", false)]
    [Arguments("DA219-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
    [Arguments("DA220-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
    [Arguments("DA221-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
    [Arguments("DA222-InvalidXPath.docx", "DA-Data.xml", true)]
    [Arguments("DA223-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
    [Arguments("DA226-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
    [Arguments("DA227-XPathErrorInPara.docx", "DA-Data.xml", true)]
    [Arguments("DA228-NoPrototypeRow.docx", "DA-Data.xml", true)]
    [Arguments("DA229-NoDataForCell.docx", "DA-Data.xml", true)]
    [Arguments("DA230-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", false)] // Clippit support multi-value XPath in table cells
    [Arguments("DA231-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
    [Arguments("DA232-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
    [Arguments("DA233-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
    [Arguments("DA234-HeaderFooter.docx", "DA-Data.xml", false)]
    [Arguments("DA235-Crashes.docx", "DA-Content-List.xml", false)]
    [Arguments("DA236-Page-Num-in-Footer.docx", "DA-Content-List.xml", false)]
    [Arguments("DA237-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
    [Arguments("DA238-SchemaErrorInConditional.docx", "DA-Data.xml", true)]
    [Arguments("DA239-RunLevelCC-Repeat.docx", "DA-Data.xml", false)]
    [Arguments("DA250-ConditionalWithRichXPath.docx", "DA250-Address.xml", false)]
    [Arguments("DA251-EnhancedTables.docx", "DA-Data.xml", false)]
    [Arguments("DA252-Table-With-Sum.docx", "DA-Data.xml", false)]
    [Arguments("DA253-Table-With-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
    [Arguments("DA254-Table-With-XPath-Sum.docx", "DA-Data.xml", false)]
    [Arguments("DA255-Table-With-XPath-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
    [Arguments("DA256-NoInvalidDocOnErrorInRun.docx", "DA-Data.xml", true)]
    [Arguments("DA257-OptionalRepeat.docx", "DA-Data.xml", false)]
    [Arguments("DA258-ContentAcceptsCharsAsXPathResult.docx", "DA-Data.xml", false)]
    [Arguments("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
    [Arguments("DA260-RunLevelRepeat.docx", "DA-Data.xml", false)]
    [Arguments("DA261-RunLevelConditional.docx", "DA-Data.xml", false)]
    [Arguments("DA262-ConditionalNotMatch.docx", "DA-Data.xml", false)]
    [Arguments("DA263-ConditionalNotMatch.docx", "DA-DataSmallCustomer.xml", false)]
    [Arguments("DA264-InvalidRunLevelRepeat.docx", "DA-Data.xml", true)]
    [Arguments("DA265-RunLevelRepeatWithWhiteSpaceBefore.docx", "DA-Data.xml", false)]
    [Arguments("DA266-RunLevelRepeat-NoData.docx", "DA-Data.xml", true)]
    [Arguments("DA267-Repeat-HorizontalAlignType.docx", "DA-Data.xml", false)]
    [Arguments("DA268-Repeat-VerticalAlignType.docx", "DA-Data.xml", false)]
    [Arguments("DA269-Repeat-InvalidAlignType.docx", "DA-Data.xml", true)]
    [Arguments("DA270-ImageSelect.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA270A-ImageSelect.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA271-ImageSelectWithRepeat.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA271A-ImageSelectWithRepeat.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA272-ImageSelectWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA272A-ImageSelectWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA273-ImageSelectInsideTextBoxWithRepeatVerticalAlign.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA273A-ImageSelectInsideTextBoxWithRepeatVerticalAlign.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA274-ImageSelectInsideTextBoxWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA274A-ImageSelectInsideTextBoxWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA275-ImageSelectWithRepeatInvalidAlign.docx", "DA-Data-WithImages.xml", true)]
    [Arguments("DA275A-ImageSelectWithRepeatInvalidAlign.docx", "DA-Data-WithImages.xml", true)]
    [Arguments("DA276-ImageSelectInsideTable.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA276A-ImageSelectInsideTable.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA277-ImageSelectMissingOrInvalidPictureContent.docx", "DA-Data-WithImages.xml", true)]
    [Arguments("DA277A-ImageSelectMissingOrInvalidPictureContent.docx", "DA-Data-WithImages.xml", true)]
    [Arguments("DA278-ImageSelect.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
    [Arguments("DA278A-ImageSelect.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
    [Arguments("DA279-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidMIMEType.xml", true)]
    [Arguments("DA279A-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidMIMEType.xml", true)]
    [Arguments("DA280-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidImageDataFormat.xml", true)]
    [Arguments("DA280A-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidImageDataFormat.xml", true)]
    [Arguments("DA281-ImageSelectExtraWhitespaceBeforeImageContent.docx", "DA-Data-WithImages.xml", true)]
    [Arguments("DA281A-ImageSelectExtraWhitespaceBeforeImageContent.docx", "DA-Data-WithImages.xml", true)]
    [Arguments("DA282-ImageSelectWithHeader.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA282A-ImageSelectWithHeader.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA282-ImageSelectWithHeader.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
    [Arguments("DA282A-ImageSelectWithHeader.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
    [Arguments("DA283-ImageSelectWithFooter.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA283A-ImageSelectWithFooter.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA284-ImageSelectWithHeaderAndFooter.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA284A-ImageSelectWithHeaderAndFooter.docx", "DA-Data-WithImages.xml", false)]
    [Arguments("DA285-ImageSelectNoParagraphFollowedAfterMetadata.docx", "DA-Data-WithImages.xml", true)]
    [Arguments("DA285A-ImageSelectNoParagraphFollowedAfterMetadata.docx", "DA-Data-WithImages.xml", true)]
    [Arguments("DA-I0038-TemplateWithMultipleXPathResults.docx", "DA-I0038-Data.xml", false)]
    public async Task DA101(string name, string data, bool err)
    {
        var afterAssembling = AssembleDocument(name, data, out var returnedTemplateError);
        var assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);

        await ValidateAsync(assembledDocx);
        await Assert.That(returnedTemplateError).IsEqualTo(err);
    }

    [Test]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-no-inline-styles.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-bold.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-strong.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-italic.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-emphasis.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-underline.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-bold-underline.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-bold-italic.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-italic-underline.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-subscript.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-superscript.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-strikethrough.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-hyperlink.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-hyperlink-bold.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-hyperlink-italic.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-hyperlink-underline.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-hyperlink-no-protocol.xml", 1, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-multi-paragraph.xml", 3, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-multi-paragraph-with-CRLF.xml", 3, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-multi-paragraph-text-only.xml", 3, false)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-invalid.xml", 0, true)]
    [Arguments("DA289-xhtml-formatting.docx", "DA289-not-well-formed.xml", 0, true)]
    public async Task DA289(string name, string data, int parasInContent, bool err)
    {
        var wmlTemplate = new WmlDocument(Path.Combine(_sourceDir.FullName, name));

        var wmlResult = AssembleDocument(name, data, out bool returnedTemplateError);
        var assembledDocx = GetOutputFile(data);
        wmlResult.SaveAs(assembledDocx.FullName);

        await ValidateAsync(assembledDocx);
        await Assert.That(returnedTemplateError).IsEqualTo(err);

        // if we are not expecting an error then verify that we have the same number of paragraphs and that
        // the paragraph properties from source and target are the same
        if (!err)
        {
            IList<XElement> sourceParas = wmlTemplate.MainDocumentPart.Element(W.body).Descendants(W.p).ToList();
            IList<XElement> targetParas = wmlResult.MainDocumentPart.Element(W.body).Descendants(W.p).ToList();
            IList<XElement> targetBreaks = wmlResult.MainDocumentPart.Element(W.body).Descendants(W.br).ToList();

            // Check we have the expected number of paragraphs
            // Expected document structure is:
            //   Heading paragraph (1 line)
            //   Empty paragraph (1 line)
            //   Escaped HTML paragraph (potential multi-line)
            //   CDATA paragraph (potential multi-line)
            await Assert.That(targetParas).HasCount(sourceParas.Count);
            int expectedBreaks = (parasInContent - 1) * 2;
            await Assert.That(targetBreaks).HasCount(expectedBreaks);

            var equalityComparer = new XNodeEqualityComparer();
            int paraOffset = 0;

            for (var i = 0; i < sourceParas.Count; i++)
            {
                var sourceProps = sourceParas[i].Element(W.pPr);
                var targetProps = targetParas[i].Element(W.pPr);

                if (sourceProps == null && targetProps == null)
                {
                    continue;
                }

                await Assert.That(equalityComparer.Equals(sourceProps, targetProps)).IsTrue();
            }
        }
    }

    [Test]
    [Arguments("DA290-xhtml-merge-run-formatting.docx", "DA290-xhtml-merge-run-formatting.xml")]
    public async Task DA290_Merge_Run_Formatting(string name, string data)
    {
        // Act
        var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
        FileInfo assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);

        // Assert - para count is expected
        List<XElement> paras = afterAssembling.MainDocumentPart.Element(W.body).Descendants(W.p).ToList();
        await Assert.That(paras).HasCount(9);

        // Assert - Paragraph 1 Styles
        XElement para = paras[0];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading1");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value).IsEqualTo("16");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value).IsEqualTo("16");
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.sz).Attribute(W.val).Value).IsEqualTo("16");
            await Assert.That(x.Element(W.szCs).Attribute(W.val).Value).IsEqualTo("16");
        }

        // Assert - Paragraph 2 Styles
        para = paras[1];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading2");
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.rStyle).Attribute(W.val).Value).IsEqualTo("Heading2Char");
        }

        // Assert - Paragraph 3 Styles
        para = paras[2];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading2CDATA");
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.color).Attribute(W.val).Value).IsEqualTo("538135");
        }

        // Assert - Paragraph 4 Styles
        para = paras[3];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading2CDATA");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.rFonts)).IsNotNull();
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value).IsEqualTo("auto");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value).IsEqualTo("22");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value).IsEqualTo("22");
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.rFonts).Attribute(W.ascii).Value).IsEqualTo("Algerian");
            await Assert.That(x.Element(W.i)).IsNotNull();
            await Assert.That(x.Element(W.iCs)).IsNotNull();
            await Assert.That(x.Element(W.u).Attribute(W.val).Value).IsEqualTo("single");
        }

        // Assert - Paragraph 5 Styles
        para = paras[4];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading2CDATA");
        await Assert
            .That(para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value)
            .IsEqualTo("C45911");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value).IsEqualTo("14");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value).IsEqualTo("14");
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.sz).Attribute(W.val).Value).IsEqualTo("14");
            await Assert.That(x.Element(W.szCs).Attribute(W.val).Value).IsEqualTo("14");
        }

        // Assert - Paragraph 6 Styles
        para = paras[5];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading2CDATA");
        await Assert
            .That(para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value)
            .IsEqualTo("C45911");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value).IsEqualTo("40");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value).IsEqualTo("40");
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.sz).Attribute(W.val).Value).IsEqualTo("40");
            await Assert.That(x.Element(W.szCs).Attribute(W.val).Value).IsEqualTo("40");
        }

        // Assert - Paragraph 7 Styles
        para = paras[6];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading2CDATA");
        await Assert
            .That(para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value)
            .IsEqualTo("00B0F0");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value).IsEqualTo("40");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value).IsEqualTo("40");
        await Assert
            .That(paras[6].Element(W.pPr).Element(W.rPr).Element(W.rFonts).Attribute(W.ascii).Value)
            .IsEqualTo("Algerian");
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.rFonts).Attribute(W.ascii).Value).IsEqualTo("Algerian");
            await Assert.That(x.Element(W.color).Attribute(W.val).Value).IsEqualTo("00B0F0");
            await Assert.That(x.Element(W.sz).Attribute(W.val).Value).IsEqualTo("40");
            await Assert.That(x.Element(W.szCs).Attribute(W.val).Value).IsEqualTo("40");
        }

        // Assert - Paragraph 8 Styles
        para = paras[7];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading2CDATA");
        await Assert
            .That(para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value)
            .IsEqualTo("C45911");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value).IsEqualTo("32");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value).IsEqualTo("32");
        await Assert
            .That(para.Element(W.pPr).Element(W.rPr).Element(W.rFonts).Attribute(W.ascii).Value)
            .IsEqualTo("Algerian");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.u).Attribute(W.val).Value).IsEqualTo("single");
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.rFonts).Attribute(W.ascii).Value).IsEqualTo("Algerian");
            await Assert.That(x.Element(W.u).Attribute(W.val).Value).IsEqualTo("single");
            await Assert.That(x.Element(W.sz).Attribute(W.val).Value).IsEqualTo("32");
            await Assert.That(x.Element(W.szCs).Attribute(W.val).Value).IsEqualTo("32");
        }

        // Assert - Paragraph 9 Styles
        para = paras[8];
        await Assert.That(para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value).IsEqualTo("Heading2CDATA");
        await Assert
            .That(para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value)
            .IsEqualTo("538135");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value).IsEqualTo("28");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value).IsEqualTo("28");
        await Assert
            .That(para.Element(W.pPr).Element(W.rPr).Element(W.rFonts).Attribute(W.ascii).Value)
            .IsEqualTo("Algerian");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.u).Attribute(W.val).Value).IsEqualTo("single");
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.i)).IsNotNull();
        await Assert.That(para.Element(W.pPr).Element(W.rPr).Element(W.iCs)).IsNotNull();
        foreach (var x in para.Descendants(W.r).Elements(W.rPr))
        {
            await Assert.That(x.Element(W.rFonts).Attribute(W.ascii).Value).IsEqualTo("Algerian");
            await Assert.That(x.Element(W.color).Attribute(W.val).Value).IsEqualTo("538135");
            await Assert.That(x.Element(W.u).Attribute(W.val).Value).IsEqualTo("single");
            await Assert.That(x.Element(W.sz).Attribute(W.val).Value).IsEqualTo("28");
            await Assert.That(x.Element(W.szCs).Attribute(W.val).Value).IsEqualTo("28");
            await Assert.That(x.Element(W.i)).IsNotNull();
            await Assert.That(x.Element(W.iCs)).IsNotNull();
        }
    }

    [Test]
    [Arguments("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
    public async Task DA259(string name, string data, bool err)
    {
        var afterAssembling = AssembleDocument(name, data, out _);
        FileInfo assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);

        var brCount = afterAssembling.MainDocumentPart.Element(W.body).Descendants(W.r).Elements(W.br).Count();

        await Assert.That(brCount).IsEqualTo(4);
    }

    [Test]
    [Arguments("DA286-DocumentTemplate-Base-Main.docx", "DA286-DocumentTemplate-Base.xml", false)]
    [Arguments("DA286-DocumentTemplate-MirroredMargins-Main.docx", "DA286-DocumentTemplate-MirroredMargins.xml", false)]
    [Arguments("DA286-DocumentTemplate-NoBreaks-Main.docx", "DA286-DocumentTemplate-NoBreaks.xml", false)]
    [Arguments("DA286-DocumentTemplate-HeaderFooter-Main.docx", "DA286-DocumentTemplate-HeaderFooter.xml", false)]
    [Arguments("DA286-Document-SolarSystem-Main.docx", "DA286-Document-SolarSystem.xml", false)]
    public async Task DA286(string templateName, string data, bool err)
    {
        var templateDocx = new FileInfo(Path.Combine(_sourceDir.FullName, templateName));
        var dataFile = new FileInfo(Path.Combine(_sourceDir.FullName, data));

        var wmlTemplate = new WmlDocument(templateDocx.FullName, true);
        var xmldata = XElement.Load(dataFile.FullName);

        // set the directory for TemplatePath attributes
        var ns = xmldata.GetDefaultNamespace();
        foreach (var ele in xmldata.XPathSelectElements("//*[@TemplatePath]"))
        {
            var templatePath = ele.Attribute(ns + "TemplatePath").Value;
            templatePath = Path.Combine(_sourceDir.FullName, templatePath);
            ele.Attribute(ns + "TemplatePath").Value = templatePath;
        }

        // set the directory for Path attributes
        foreach (var ele in xmldata.XPathSelectElements("//*[@Path]"))
        {
            var path = ele.Attribute(ns + "Path").Value;
            path = Path.Combine(_sourceDir.FullName, path);
            ele.Attribute(ns + "Path").Value = path;
        }

        var afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out bool templateError);
        var assembledDocx = new FileInfo(
            Path.Combine(TempDir, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx"))
        );
        afterAssembling.SaveAs(assembledDocx.FullName);

        await ValidateAsync(assembledDocx);

        await Assert.That(templateError).IsEqualTo(err);
    }

    [Test]
    [Arguments("DA024-TrackedRevisions.docx", "DA-Data.xml")]
    public async Task DA102_Throws(string name, string data)
    {
        var templateDocx = new FileInfo(Path.Combine(_sourceDir.FullName, name));
        var dataFile = new FileInfo(Path.Combine(_sourceDir.FullName, data));

        var wmlTemplate = new WmlDocument(templateDocx.FullName);
        var xmldata = XElement.Load(dataFile.FullName);

        WmlDocument afterAssembling;
        await Assert
            .That(() =>
            {
                afterAssembling = DocumentAssembler.AssembleDocument(wmlTemplate, xmldata, out _);
            })
            .Throws<OpenXmlPowerToolsException>();
    }

    [Test]
    [Arguments("DA-TemplateMaior.docx", "DA-templateMaior.xml", false)]
    public async Task DATemplateMaior(string name, string data, bool err)
    {
        var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
        FileInfo assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);

        var descendants = afterAssembling.MainDocumentPart.Value;

        await Assert.That(descendants.Contains(">")).IsFalse().Because("Found > on text");
    }

    [Test]
    [Arguments("DA-xmlerror.docx", "DA-xmlerror.xml")]
    public async Task DAXmlError(string name, string data)
    {
        var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
        FileInfo assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);
    }

    [Test]
    [Arguments("DA025-TemplateDocument.docx", "DA-Data.xml", false)]
    public async Task DA103_UseXmlDocument(string name, string data, bool err)
    {
        var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
        FileInfo assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);

        await ValidateAsync(assembledDocx);
        await Assert.That(returnedTemplateError).IsEqualTo(err);
    }

    [Test]
    [Arguments("DA-Ampersand+LF-Issue.docx", "DA-Ampersand+LF-Issue.xml", false)]
    [Arguments("DA-Ampersand+LF-Issue-With-Controls.docx", "DA-Ampersand+LF-Issue.xml", false)]
    public async Task DA_Ampersands_And_LineFeeds(string name, string data, bool err)
    {
        var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
        FileInfo assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);

        // Assert - no errors
        await ValidateAsync(assembledDocx);
        await Assert.That(returnedTemplateError).IsEqualTo(err);

        // Assert - tables is present and correct
        XElement table = afterAssembling.MainDocumentPart.Descendants(W.tbl).SingleOrDefault();
        await Assert.That(table).IsNotNull();

        // Assert - the second table cell of each table has one paragraph
        IEnumerable<XElement> paras = table.Descendants(W.tc).ElementAt(1).Elements(W.p);
        await Assert.That(paras).HasSingleItem();

        // Assert - first table paragraph has 2 soft breaks
        await Assert.That(paras.ElementAt(0).Elements(W.r)).HasCount(5);
        await Assert.That(paras.ElementAt(0).Elements(W.r).Elements(W.br)).HasCount(2);
    }

    [Test]
    [Arguments("DA-Tabs-In-Text.docx", "DA-Tabs-In-Text.xml", false)]
    [Arguments("DA-Tabs-In-Text-With-Controls.docx", "DA-Tabs-In-Text.xml", false)]
    public async Task DA_Tabs_In_Text(string name, string data, bool err)
    {
        var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
        FileInfo assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);

        // Assert - no errors
        await ValidateAsync(assembledDocx);
        await Assert.That(returnedTemplateError).IsEqualTo(err);

        // Assert - we have four paragraphs
        IEnumerable<XElement> paras = afterAssembling.MainDocumentPart.Descendants(W.p);
        await Assert.That(paras).HasCount(4);

        // Assert - first paragraph has 0 tabs
        await Assert.That(paras.ElementAt(0).Descendants(W.tab)).IsEmpty();

        // Assert - second paragraph has a tab in the first run
        await Assert.That(paras.ElementAt(1).Elements(W.r).First().Elements(W.tab)).HasSingleItem();

        // Assert - third paragraph has a tab in the last run
        await Assert.That(paras.ElementAt(2).Elements(W.r).Last().Elements(W.tab)).HasSingleItem();

        // Assert - fourth paragraph has a tab but not in the first or last run
        await Assert.That(paras.ElementAt(3).Descendants(W.tab)).HasSingleItem();
        await Assert.That(paras.ElementAt(3).Elements(W.r).First().Elements(W.tab)).IsEmpty();
        await Assert.That(paras.ElementAt(3).Elements(W.r).Last().Elements(W.tab)).IsEmpty();
    }

    [Test]
    [Arguments("DA-Issue-95-Template.docx", "DA-Issue-95-Data.xml", false)]
    public async Task DA_Issue_95_Repro(string name, string data, bool err)
    {
        var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
        FileInfo assembledDocx = GetOutputFile(name);
        afterAssembling.SaveAs(assembledDocx.FullName);

        // Assert - no errors
        ValidateAsync(assembledDocx);
        await Assert.That(returnedTemplateError).IsEqualTo(err);

        // Assert - tables are present and correct
        IEnumerable<XElement> tables = afterAssembling.MainDocumentPart.Descendants(W.tbl);
        await Assert.That(tables).HasCount(4);

        // Assert - the second table cell of each table has one paragraph
        List<XElement> paras = new List<XElement>();
        foreach (XElement table in tables)
        {
            paras.AddRange(table.Descendants(W.tc).ElementAt(1).Elements(W.p));
        }

        await Assert.That(tables).HasCount(paras.Count);

        // Assert - first tables paragraph has 4 soft breaks
        await Assert.That(paras.ElementAt(0).Elements(W.r)).HasCount(7);
        await Assert.That(paras.ElementAt(0).Elements(W.r).Elements(W.br)).HasCount(4);

        // Assert - second tables paragraph has 1 soft breaks
        await Assert.That(paras.ElementAt(1).Elements(W.r)).HasCount(3);
        await Assert.That(paras.ElementAt(1).Elements(W.r).Elements(W.br)).HasSingleItem();

        // Assert - third tables paragraph has 2 soft breaks
        await Assert.That(paras.ElementAt(2).Elements(W.r)).HasCount(5);
        await Assert.That(paras.ElementAt(2).Elements(W.r).Elements(W.br)).HasCount(2);

        // Assert - fourth tables paragraph has 1 soft breaks and two tabs
        await Assert.That(paras.ElementAt(3).Elements(W.r)).HasCount(5);
        await Assert.That(paras.ElementAt(3).Elements(W.r).Elements(W.br)).HasSingleItem();
        await Assert.That(paras.ElementAt(3).Elements(W.r).Elements(W.tab)).HasCount(2);
    }

    private async Task ValidateAsync(FileInfo fi)
    {
        using var wDoc = WordprocessingDocument.Open(fi.FullName, false);
        await Validate(wDoc, s_expectedErrors);
    }

    private WmlDocument AssembleDocument(string templateFilename, string xmlFilename, out bool templateError)
    {
        var templateDocx = new FileInfo(Path.Combine(_sourceDir.FullName, templateFilename));
        var dataFile = new FileInfo(Path.Combine(_sourceDir.FullName, xmlFilename));

        var wmlTemplate = new WmlDocument(templateDocx.FullName);
        var xmlData = new XmlDocument();
        xmlData.Load(dataFile.FullName);

        return DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out templateError);
    }

    private FileInfo GetOutputFile(string fileName)
    {
        return new FileInfo(
            Path.Combine(TempDir, fileName.Replace(Path.GetExtension(fileName), "-processed-by-DocumentAssembler.docx"))
        );
    }

    private static readonly List<string> s_expectedErrors = new()
    {
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:evenVBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRow' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowFirstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:firstRowLastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRow' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowFirstColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:lastRowLastColumn' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:noVBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddHBand' attribute is not declared.",
        "The 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:oddVBand' attribute is not declared.",
        "The 'http://schemas.microsoft.com/office/word/2012/wordml:restartNumberingAfterBreak' attribute is not declared.",
        "The 'http://schemas.microsoft.com/office/word/2016/wordml/cid:durableId' attribute is not declared.",
        "Attribute 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:val' should have unique value. Its current value",
        "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:bCs'.",
        "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:rFonts'.",
        "The element has unexpected child element 'http://schemas.openxmlformats.org/wordprocessingml/2006/main:kern'.",
    };
}
