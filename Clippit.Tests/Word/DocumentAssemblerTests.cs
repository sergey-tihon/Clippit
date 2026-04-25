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
    [Arguments("DA291-Table-Optional-NoData.docx", "DA-Data.xml", false)]
    [Arguments("DA291A-Table-Optional-NoData.docx", "DA-Data.xml", false)]
    [Arguments("DA292-Table-Optional-WithData.docx", "DA-Data.xml", false)]
    [Arguments("DA292A-Table-Optional-WithData.docx", "DA-Data.xml", false)]
    public async Task DA101(string name, string data, bool err)
    {
        var afterAssembling = AssembleDocument(name, data, out var returnedTemplateError);
        var assembledDocx = GetOutputFile(name, data);
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
            await Assert.That(targetParas).Count().IsEqualTo(sourceParas.Count);
            int expectedBreaks = (parasInContent - 1) * 2;
            await Assert.That(targetBreaks).Count().IsEqualTo(expectedBreaks);

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
        await Assert.That(paras).Count().IsEqualTo(9);

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
        await Assert.That(paras.ElementAt(0).Elements(W.r)).Count().IsEqualTo(5);
        await Assert.That(paras.ElementAt(0).Elements(W.r).Elements(W.br)).Count().IsEqualTo(2);
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
        await Assert.That(paras).Count().IsEqualTo(4);

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
        await Assert.That(tables).Count().IsEqualTo(4);

        // Assert - the second table cell of each table has one paragraph
        List<XElement> paras = [];
        foreach (XElement table in tables)
        {
            paras.AddRange(table.Descendants(W.tc).ElementAt(1).Elements(W.p));
        }

        await Assert.That(tables).Count().IsEqualTo(paras.Count);

        // Assert - first tables paragraph has 4 soft breaks
        await Assert.That(paras.ElementAt(0).Elements(W.r)).Count().IsEqualTo(7);
        await Assert.That(paras.ElementAt(0).Elements(W.r).Elements(W.br)).Count().IsEqualTo(4);

        // Assert - second tables paragraph has 1 soft breaks
        await Assert.That(paras.ElementAt(1).Elements(W.r)).Count().IsEqualTo(3);
        await Assert.That(paras.ElementAt(1).Elements(W.r).Elements(W.br)).HasSingleItem();

        // Assert - third tables paragraph has 2 soft breaks
        await Assert.That(paras.ElementAt(2).Elements(W.r)).Count().IsEqualTo(5);
        await Assert.That(paras.ElementAt(2).Elements(W.r).Elements(W.br)).Count().IsEqualTo(2);

        // Assert - fourth tables paragraph has 1 soft breaks and two tabs
        await Assert.That(paras.ElementAt(3).Elements(W.r)).Count().IsEqualTo(5);
        await Assert.That(paras.ElementAt(3).Elements(W.r).Elements(W.br)).HasSingleItem();
        await Assert.That(paras.ElementAt(3).Elements(W.r).Elements(W.tab)).Count().IsEqualTo(2);
    }

    /// <summary>
    /// Regression test for issue #62: <c>v:textpath/@string</c> attributes that contain a
    /// <c>&lt;#&lt;Content Select="..."/&gt;#&gt;</c> template directive were never substituted
    /// because <see cref="DocumentAssembler"/> only scanned text nodes, not XML attributes.
    /// </summary>
    [Test]
    public async Task DA_VmlTextpath_WatermarkSubstitution()
    {
        // Arrange — build a minimal DOCX in memory whose body contains a paragraph with a
        // VML shape carrying the watermark template directive in v:textpath/@string.
        const string watermarkText = "CONFIDENTIAL";
        const string xmlDataStr = "<Data><WaterMark>CONFIDENTIAL</WaterMark></Data>";

        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace vml = "urn:schemas-microsoft-com:vml";

        // The paragraph reproduces the real watermark XML produced by Word:
        //   <w:p><w:pict><v:shape><v:textpath string="<#<Content Select=&quot;./WaterMark&quot;/>#>"/></v:shape></w:pict></w:p>
        var bodyXml = new XElement(
            W.body,
            new XElement(
                W.p,
                new XElement(
                    W.pict,
                    new XElement(
                        vml + "shape",
                        new XElement(
                            vml + "textpath",
                            new XAttribute("string", "<#<Content Select=\"./WaterMark\"/>#>"),
                            new XAttribute("style", "font-family:Calibri;font-size:1pt")
                        )
                    )
                )
            ),
            new XElement(W.sectPr)
        );

        byte[] docxBytes;
        using (var ms = new MemoryStream())
        {
            using (
                var wordDoc = WordprocessingDocument.Create(
                    ms,
                    DocumentFormat.OpenXml.WordprocessingDocumentType.Document
                )
            )
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.PutXDocument(new XDocument(new XElement(W.document, bodyXml)));
            }
            docxBytes = ms.ToArray();
        }

        var wmlTemplate = new WmlDocument("watermark-template.docx", docxBytes);
        var xmlData = XElement.Parse(xmlDataStr);

        // Act
        var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

        // Assert
        await Assert.That(hasError).IsFalse();

        using var resultStream = new MemoryStream(result.DocumentByteArray);
        using var resultDoc = WordprocessingDocument.Open(resultStream, false);
        var resultBody = resultDoc.MainDocumentPart.GetXDocument().Root?.Element(W.body);
        var textpathAttr = resultBody?.Descendants(vml + "textpath").FirstOrDefault()?.Attribute("string");
        await Assert.That(textpathAttr).IsNotNull();
        await Assert.That(textpathAttr!.Value).IsEqualTo(watermarkText);
    }

    /// <summary>
    /// Negative tests for <c>v:textpath/@string</c> substitution: verifies that malformed
    /// directives set <c>hasError = true</c> and leave a descriptive placeholder in the attribute.
    /// </summary>
    [Test]
    [Arguments("<#not-valid-xml#>", "<Data><WaterMark>CONFIDENTIAL</WaterMark></Data>")]
    [Arguments("<#<Content Select=\"./Missing\"/>#>", "<Data><WaterMark>CONFIDENTIAL</WaterMark></Data>")]
    [Arguments("<#<Content Select=\"./Item\"/>#>", "<Data><Item>A</Item><Item>B</Item></Data>")]
    [Arguments("<#<Content Optional=\"true\"/>#>", "<Data><WaterMark>CONFIDENTIAL</WaterMark></Data>")]
    public async Task DA_VmlTextpath_ErrorDirective_SetsHasError(string stringAttrValue, string xmlDataStr)
    {
        XNamespace vml = "urn:schemas-microsoft-com:vml";

        var bodyXml = new XElement(
            W.body,
            new XElement(
                W.p,
                new XElement(
                    W.pict,
                    new XElement(
                        vml + "shape",
                        new XElement(
                            vml + "textpath",
                            new XAttribute("string", stringAttrValue),
                            new XAttribute("style", "font-family:Calibri;font-size:1pt")
                        )
                    )
                )
            ),
            new XElement(W.sectPr)
        );

        byte[] docxBytes;
        using (var ms = new MemoryStream())
        {
            using (
                var wordDoc = WordprocessingDocument.Create(
                    ms,
                    DocumentFormat.OpenXml.WordprocessingDocumentType.Document
                )
            )
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.PutXDocument(new XDocument(new XElement(W.document, bodyXml)));
            }
            docxBytes = ms.ToArray();
        }

        var wmlTemplate = new WmlDocument("watermark-error-template.docx", docxBytes);
        var xmlData = XElement.Parse(xmlDataStr);

        // Act
        var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

        // Assert — any directive error must surface via hasError
        await Assert.That(hasError).IsTrue();

        // The attribute value should contain the error placeholder, not the original directive.
        using var resultStream = new MemoryStream(result.DocumentByteArray);
        using var resultDoc = WordprocessingDocument.Open(resultStream, false);

        var resultBody = resultDoc.MainDocumentPart!.GetXDocument().Root?.Element(W.body);
        var textpathAttr = resultBody?.Descendants(vml + "textpath").FirstOrDefault()?.Attribute("string");

        await Assert.That(textpathAttr).IsNotNull();
        await Assert.That(textpathAttr!.Value).Contains("[Template error:");
    }

    [Test]
    public async Task DA291_Image_FitWithin_SmallImage_KeepsOriginalSize()
    {
        // Template has a 200x200 px equivalent placeholder (1905000 x 1905000 EMUs)
        // Input image is 50x50 px → fits within bounds → output should be 50x50 px in EMUs
        const int pixelInEmu = 914400 / 96; // 9525
        const long templateSizeEmu = 200L * pixelInEmu; // 1905000
        const int imageWidth = 50,
            imageHeight = 50;

        var templateBytes = BuildImageFitWithinTemplate(templateSizeEmu, templateSizeEmu);
        var imageBytes = BuildTestPng(imageWidth, imageHeight);
        var base64 = Convert.ToBase64String(imageBytes);
        var xmlData = XElement.Parse($"<Root><Photo>data:image/png;base64,{base64}</Photo></Root>");

        var wml = new WmlDocument("test.docx", templateBytes);
        var result = DocumentAssembler.AssembleDocument(wml, xmlData, out var hasError);

        await Assert.That(hasError).IsFalse();

        using var ms = new MemoryStream(result.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var xDoc = doc.MainDocumentPart!.GetXDocument();
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var extent = xDoc.Descendants(wp + "extent").FirstOrDefault();
        var xfrmExt = xDoc.Descendants(a + "xfrm").Elements(a + "ext").FirstOrDefault();

        await Assert.That(extent).IsNotNull();
        await Assert.That(xfrmExt).IsNotNull();
        var expectedEmu = (long)imageWidth * pixelInEmu; // 476250
        await Assert.That(long.Parse(extent!.Attribute("cx")!.Value)).IsEqualTo(expectedEmu);
        await Assert.That(long.Parse(extent!.Attribute("cy")!.Value)).IsEqualTo(expectedEmu);
        await Assert.That(long.Parse(xfrmExt!.Attribute("cx")!.Value)).IsEqualTo(expectedEmu);
        await Assert.That(long.Parse(xfrmExt!.Attribute("cy")!.Value)).IsEqualTo(expectedEmu);
    }

    [Test]
    public async Task DA292_Image_FitWithin_LargeImage_ScalesDownProportionally()
    {
        // Template has a 200x200 px equivalent placeholder (1905000 x 1905000 EMUs)
        // Input image is 400x200 px (wider than tall) → scale down to fit
        // Expected: scale = min(200/400, 200/200) = 0.5 → finalCx=200px, finalCy=100px in EMUs
        const int pixelInEmu = 914400 / 96; // 9525
        const long templateSizeEmu = 200L * pixelInEmu; // 1905000
        const int imageWidth = 400,
            imageHeight = 200;

        var templateBytes = BuildImageFitWithinTemplate(templateSizeEmu, templateSizeEmu);
        var imageBytes = BuildTestPng(imageWidth, imageHeight);
        var base64 = Convert.ToBase64String(imageBytes);
        var xmlData = XElement.Parse($"<Root><Photo>data:image/png;base64,{base64}</Photo></Root>");

        var wml = new WmlDocument("test.docx", templateBytes);
        var result = DocumentAssembler.AssembleDocument(wml, xmlData, out var hasError);

        await Assert.That(hasError).IsFalse();

        using var ms = new MemoryStream(result.DocumentByteArray);
        using var doc = WordprocessingDocument.Open(ms, false);
        var xDoc = doc.MainDocumentPart!.GetXDocument();
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var extent = xDoc.Descendants(wp + "extent").FirstOrDefault();
        var xfrmExt = xDoc.Descendants(a + "xfrm").Elements(a + "ext").FirstOrDefault();

        await Assert.That(extent).IsNotNull();
        await Assert.That(xfrmExt).IsNotNull();
        // scale = 200/400 = 0.5; finalCx = 400*0.5*9525 = 1905000; finalCy = 200*0.5*9525 = 952500
        long expectedCx = templateSizeEmu; // 1905000 — scales to exactly the max width
        long expectedCy = (long)(imageHeight * 0.5 * pixelInEmu); // 952500
        await Assert.That(long.Parse(extent!.Attribute("cx")!.Value)).IsEqualTo(expectedCx);
        await Assert.That(long.Parse(extent!.Attribute("cy")!.Value)).IsEqualTo(expectedCy);
        await Assert.That(long.Parse(xfrmExt!.Attribute("cx")!.Value)).IsEqualTo(expectedCx);
        await Assert.That(long.Parse(xfrmExt!.Attribute("cy")!.Value)).IsEqualTo(expectedCy);
    }

    /// <summary>
    /// Creates a minimal DOCX template with an Image directive using FitWithin="true"
    /// and an image placeholder with the specified extent dimensions.
    /// </summary>
    private static byte[] BuildImageFitWithinTemplate(long extentCx, long extentCy)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XNamespace pic = "http://schemas.openxmlformats.org/drawingml/2006/picture";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        // The image directive content control text (split across runs so concatenation = "<Image .../>")
        var directiveSdt = new XElement(
            w + "sdt",
            new XElement(w + "sdtPr", new XElement(w + "text")),
            new XElement(w + "sdtEndPr"),
            new XElement(
                w + "sdtContent",
                new XElement(w + "r", new XElement(w + "t", "<")),
                new XElement(w + "r", new XElement(w + "t", "Image Select=\"./Photo\" FitWithin=\"true\"")),
                new XElement(w + "r", new XElement(w + "t", "/>"))
            )
        );

        // The image placeholder content control with the drawing element
        var imagePlaceholderSdt = new XElement(
            w + "sdt",
            new XElement(w + "sdtPr", new XElement(w + "picture")),
            new XElement(w + "sdtEndPr"),
            new XElement(
                w + "sdtContent",
                new XElement(
                    w + "p",
                    new XElement(
                        w + "r",
                        new XElement(
                            w + "drawing",
                            new XElement(
                                wp + "inline",
                                new XAttribute("distT", "0"),
                                new XAttribute("distB", "0"),
                                new XAttribute("distL", "0"),
                                new XAttribute("distR", "0"),
                                new XElement(
                                    wp + "extent",
                                    new XAttribute("cx", extentCx),
                                    new XAttribute("cy", extentCy)
                                ),
                                new XElement(
                                    wp + "docPr",
                                    new XAttribute("id", "1"),
                                    new XAttribute("name", "Image 1")
                                ),
                                new XElement(
                                    wp + "cNvGraphicFramePr",
                                    new XElement(
                                        a + "graphicFrameLocks",
                                        new XAttribute(
                                            XNamespace.Xmlns + "a",
                                            "http://schemas.openxmlformats.org/drawingml/2006/main"
                                        ),
                                        new XAttribute("noChangeAspect", "1")
                                    )
                                ),
                                new XElement(
                                    a + "graphic",
                                    new XAttribute(
                                        XNamespace.Xmlns + "a",
                                        "http://schemas.openxmlformats.org/drawingml/2006/main"
                                    ),
                                    new XElement(
                                        a + "graphicData",
                                        new XAttribute(
                                            "uri",
                                            "http://schemas.openxmlformats.org/drawingml/2006/picture"
                                        ),
                                        new XElement(
                                            pic + "pic",
                                            new XAttribute(
                                                XNamespace.Xmlns + "pic",
                                                "http://schemas.openxmlformats.org/drawingml/2006/picture"
                                            ),
                                            new XElement(
                                                pic + "nvPicPr",
                                                new XElement(
                                                    pic + "cNvPr",
                                                    new XAttribute("id", "0"),
                                                    new XAttribute("name", "Image 1")
                                                ),
                                                new XElement(pic + "cNvPicPr")
                                            ),
                                            new XElement(
                                                pic + "blipFill",
                                                new XElement(a + "blip", new XAttribute(r + "embed", "rId1")),
                                                new XElement(a + "stretch", new XElement(a + "fillRect"))
                                            ),
                                            new XElement(
                                                pic + "spPr",
                                                new XElement(
                                                    a + "xfrm",
                                                    new XElement(
                                                        a + "off",
                                                        new XAttribute("x", "0"),
                                                        new XAttribute("y", "0")
                                                    ),
                                                    new XElement(
                                                        a + "ext",
                                                        new XAttribute("cx", extentCx),
                                                        new XAttribute("cy", extentCy)
                                                    )
                                                ),
                                                new XElement(
                                                    a + "prstGeom",
                                                    new XAttribute("prst", "rect"),
                                                    new XElement(a + "avLst")
                                                )
                                            )
                                        )
                                    )
                                )
                            )
                        )
                    )
                )
            )
        );

        var bodyXml = new XElement(
            w + "body",
            new XElement(w + "p", directiveSdt),
            imagePlaceholderSdt,
            new XElement(w + "sectPr")
        );

        using var ms = new MemoryStream();
        using (
            var wordDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
        )
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.PutXDocument(new XDocument(new XElement(w + "document", bodyXml)));

            // Add a 1x1 placeholder image so the template is valid
            var imagePart = mainPart.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Png, "rId1");
            imagePart.FeedData(new MemoryStream(BuildTestPng(1, 1)));
        }

        return ms.ToArray();
    }

    /// <summary>Creates a minimal solid-color PNG of the given dimensions.</summary>
    private static byte[] BuildTestPng(int width, int height)
    {
        using var newImage = new SixLabors.ImageSharp.Image<SixLabors.ImageSharp.PixelFormats.Rgba32>(width, height);
        newImage.ProcessPixelRows(accessor =>
        {
            for (var y = 0; y < accessor.Height; y++)
            {
                var row = accessor.GetRowSpan(y);
                for (var x = 0; x < row.Length; x++)
                    row[x] = new SixLabors.ImageSharp.PixelFormats.Rgba32(100, 149, 237, 255); // cornflower blue
            }
        });
        using var outMs = new MemoryStream();
        newImage.Save(outMs, new SixLabors.ImageSharp.Formats.Png.PngEncoder());
        return outMs.ToArray();
    }

    /// <summary>
    /// Regression test for issue #85: <c>&lt;Table Select="..." Optional="true|1"/&gt;</c>
    /// should suppress the "Table Select returned no data" error and remove the table
    /// when the XPath expression returns no matching elements. Both XSD boolean forms
    /// (<c>true</c> and <c>1</c>) are accepted.
    /// </summary>
    [Test]
    [Arguments("true", true)]
    [Arguments("1", true)]
    [Arguments("true", false)]
    [Arguments("1", false)]
    public async Task DA_Table_Optional_NoDataRemovesTable(string optionalValue, bool useSdt)
    {
        var directiveParagraph = new XElement(
            W.p,
            new XElement(W.r, new XElement(W.t, $@"<# <Table Select=""Orders"" Optional=""{optionalValue}"" /> #>"))
        );

        if (useSdt)
        {
            directiveParagraph = GetSdtFromMetadata(directiveParagraph);
        }

        var tableXml = new XElement(
            W.tbl,
            new XElement(W.tblPr),
            new XElement(W.tr, new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "Header"))))),
            new XElement(W.tr, new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "Row")))))
        );

        var bodyXml = new XElement(W.body, directiveParagraph, tableXml, new XElement(W.sectPr));

        byte[] docxBytes;
        using (var ms = new MemoryStream())
        {
            using (
                var wordDoc = WordprocessingDocument.Create(
                    ms,
                    DocumentFormat.OpenXml.WordprocessingDocumentType.Document
                )
            )
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.PutXDocument(new XDocument(new XElement(W.document, bodyXml)));
            }
            docxBytes = ms.ToArray();
        }

        var wmlTemplate = new WmlDocument($"optional-{optionalValue}-table-template.docx", docxBytes);
        var xmlData = XElement.Parse("<Data/>");

        var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

        await Assert.That(hasError).IsFalse();

        using var resultStream = new MemoryStream(result.DocumentByteArray);
        using var resultDoc = WordprocessingDocument.Open(resultStream, false);
        await Validate(resultDoc, s_expectedErrors);

        var resultBody = resultDoc.MainDocumentPart!.GetXDocument().Root?.Element(W.body);
        var tables = resultBody?.Elements(W.tbl).ToList();
        await Assert.That(tables).IsEmpty();
    }

    /// <summary>
    /// Verifies that <c>&lt;Table Select="..." Optional="true"/&gt;</c> populates the table
    /// normally when the XPath expression returns matching elements.
    /// </summary>
    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task DA_Table_Optional_WithDataPopulatesTable(bool useSdt)
    {
        var directiveParagraph = new XElement(
            W.p,
            new XElement(W.r, new XElement(W.t, @"<# <Table Select=""Items/Item"" Optional=""true"" /> #>"))
        );

        if (useSdt)
        {
            directiveParagraph = GetSdtFromMetadata(directiveParagraph);
        }

        var tableXml = new XElement(
            W.tbl,
            new XElement(W.tblPr),
            new XElement(W.tblGrid, new XElement(W.gridCol, new XAttribute(W._w, "9216"))),
            // Header row
            new XElement(W.tr, new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "Name"))))),
            // Prototype row: cells contain XPath expressions as raw text
            new XElement(W.tr, new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "./Name")))))
        );

        var bodyXml = new XElement(W.body, directiveParagraph, tableXml, new XElement(W.sectPr));

        byte[] docxBytes;
        using (var ms = new MemoryStream())
        {
            using (
                var wordDoc = WordprocessingDocument.Create(
                    ms,
                    DocumentFormat.OpenXml.WordprocessingDocumentType.Document
                )
            )
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.PutXDocument(new XDocument(new XElement(W.document, bodyXml)));
            }
            docxBytes = ms.ToArray();
        }

        var wmlTemplate = new WmlDocument("optional-table-with-data.docx", docxBytes);
        var xmlData = XElement.Parse(
            "<Data><Items><Item><Name>Apple</Name></Item><Item><Name>Banana</Name></Item></Items></Data>"
        );

        var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

        await Assert.That(hasError).IsFalse();

        using var resultStream = new MemoryStream(result.DocumentByteArray);
        using var resultDoc = WordprocessingDocument.Open(resultStream, false);
        await Validate(resultDoc, s_expectedErrors);

        var resultBody = resultDoc.MainDocumentPart!.GetXDocument().Root?.Element(W.body);
        var tables = resultBody?.Elements(W.tbl).ToList();
        await Assert.That(tables).IsNotEmpty();

        var dataRows = tables![0].Elements(W.tr).Skip(1).ToList(); // skip header row
        await Assert.That(dataRows).Count().IsEqualTo(2);

        var rowText = dataRows
            .SelectMany(r => r.Descendants(W.t))
            .Select(t => (string)t)
            .Aggregate(string.Empty, string.Concat);
        await Assert.That(rowText).Contains("Apple");
        await Assert.That(rowText).Contains("Banana");
    }

    /// <summary>
    /// Verifies that <c>&lt;Table Select="..."/&gt;</c> without <c>Optional="true"</c>
    /// still returns an error when no data is found (existing behaviour preserved).
    /// </summary>
    [Test]
    [Arguments(true)]
    [Arguments(false)]
    public async Task DA_Table_NoOptional_NoDataReturnsError(bool useSdt)
    {
        var directiveParagraph = new XElement(
            W.p,
            new XElement(W.r, new XElement(W.t, @"<# <Table Select=""Orders"" /> #>"))
        );

        if (useSdt)
        {
            directiveParagraph = GetSdtFromMetadata(directiveParagraph);
        }

        var tableXml = new XElement(
            W.tbl,
            new XElement(W.tblPr),
            new XElement(W.tr, new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "Header"))))),
            new XElement(W.tr, new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "Row")))))
        );

        var bodyXml = new XElement(W.body, directiveParagraph, tableXml, new XElement(W.sectPr));

        byte[] docxBytes;
        using (var ms = new MemoryStream())
        {
            using (
                var wordDoc = WordprocessingDocument.Create(
                    ms,
                    DocumentFormat.OpenXml.WordprocessingDocumentType.Document
                )
            )
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.PutXDocument(new XDocument(new XElement(W.document, bodyXml)));
            }
            docxBytes = ms.ToArray();
        }

        var wmlTemplate = new WmlDocument("non-optional-table-template.docx", docxBytes);
        var xmlData = XElement.Parse("<Data/>");

        var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

        await Assert.That(hasError).IsTrue();

        using var resultStream = new MemoryStream(result.DocumentByteArray);
        using var resultDoc = WordprocessingDocument.Open(resultStream, false);
        var documentText = resultDoc
            .MainDocumentPart!.GetXDocument()
            .Descendants(W.t)
            .Select(t => (string)t)
            .Aggregate(string.Empty, string.Concat);
        await Assert.That(documentText).Contains("Table Select returned no data");
    }

    /// <summary>
    /// Verifies that an invalid <c>Optional</c> attribute value produces a template error
    /// instead of throwing a <see cref="FormatException"/>.
    /// </summary>
    [Test]
    public async Task DA_Table_InvalidOptionalValueReturnsError()
    {
        // Build a template that contains a metadata element with an invalid Optional value,
        // placed directly at the body level (bypassing ValidatePerSchema) to simulate
        // a template where metadata elements already exist without going through text parsing.
        // PA.Table = "Table", PA.Select = "Select", PA.Optional = "Optional" (no namespace).
        var tableDirective = new XElement(
            "Table",
            new XAttribute("Select", "Orders"),
            new XAttribute("Optional", "yes")
        ); // invalid XSD boolean — not true/false/1/0

        var tableXml = new XElement(
            W.tbl,
            new XElement(W.tblPr),
            new XElement(W.tr, new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "Header"))))),
            new XElement(W.tr, new XElement(W.tc, new XElement(W.p, new XElement(W.r, new XElement(W.t, "Row")))))
        );

        // Place the Table directive and the table at body level so NormalizeTablesRepeatAndConditional
        // can move the w:tbl inside the directive, as it does in a normal assembled document.
        var bodyXml = new XElement(W.body, tableDirective, tableXml, new XElement(W.sectPr));

        byte[] docxBytes;
        using (var ms = new MemoryStream())
        {
            using (
                var wordDoc = WordprocessingDocument.Create(
                    ms,
                    DocumentFormat.OpenXml.WordprocessingDocumentType.Document
                )
            )
            {
                var mainPart = wordDoc.AddMainDocumentPart();
                mainPart.PutXDocument(new XDocument(new XElement(W.document, bodyXml)));
            }
            docxBytes = ms.ToArray();
        }

        var wmlTemplate = new WmlDocument("invalid-optional-table-template.docx", docxBytes);
        var xmlData = XElement.Parse("<Data/>");

        var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

        await Assert.That(hasError).IsTrue();

        using var resultStream = new MemoryStream(result.DocumentByteArray);
        using var resultDoc = WordprocessingDocument.Open(resultStream, false);
        var documentText = resultDoc
            .MainDocumentPart!.GetXDocument()
            .Descendants(W.t)
            .Select(t => (string)t)
            .Aggregate(string.Empty, string.Concat);
        await Assert.That(documentText).Contains("Invalid value for Optional attribute");
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

    private static XElement GetSdtFromMetadata(XElement element)
    {
        var Wt = element.Descendants(W.t).First();
        var text = Wt.Value?.Trim() ?? string.Empty;
        if (text.StartsWith("<#"))
        {
            text = text.Substring(2);
        }
        if (text.EndsWith("#>"))
        {
            text = text.Substring(0, text.Length - 2);
        }
        Wt.Value = text.Trim();

        return new XElement(W.sdt, new XElement(W.sdtContent, element));
    }

    private FileInfo GetOutputFile(string templateName, string dataName = null)
    {
        var baseName = Path.GetFileNameWithoutExtension(templateName);
        if (!string.IsNullOrEmpty(dataName))
        {
            var dataBaseName = Path.GetFileNameWithoutExtension(dataName);
            baseName = $"{baseName}_{dataBaseName}";
        }
        return new FileInfo(Path.Combine(TempDir, $"{baseName}-processed-by-DocumentAssembler.docx"));
    }

    private static readonly List<string> s_expectedErrors =
    [
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
    ];

    // ── Custom handler tests ──────────────────────────────────────────────────────────────

    private static byte[] BuildMinimalDocxWithDirective(string directiveXml)
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var bodyXml = new XElement(
            w + "body",
            new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", directiveXml))),
            new XElement(w + "sectPr")
        );
        using var ms = new MemoryStream();
        using (
            var wordDoc = WordprocessingDocument.Create(ms, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)
        )
        {
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.PutXDocument(new XDocument(new XElement(w + "document", bodyXml)));
        }
        return ms.ToArray();
    }

    /// <summary>DA500 — a registered custom handler is invoked and its return value replaces the directive.</summary>
    [Test]
    public async Task DA500_CustomHandler_RegisterAndInvoke()
    {
        const string elementName = "DA500_Greeting";
        try
        {
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

            DocumentAssembler.RegisterCustomHandler(
                elementName,
                schemaXsd: null,
                handler: (directive, data, part) =>
                {
                    var name = (string)data.Element("Name") ?? "World";
                    return new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", $"Hello, {name}!")));
                }
            );

            var directive = $"<#<{elementName}/>#>";
            var docxBytes = BuildMinimalDocxWithDirective(directive);
            var wmlTemplate = new WmlDocument("test500.docx", docxBytes);
            var xmlData = XElement.Parse("<Root><Name>Alice</Name></Root>");

            var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

            await Assert.That(hasError).IsFalse();

            using var resultStream = new MemoryStream(result.DocumentByteArray);
            using var resultDoc = WordprocessingDocument.Open(resultStream, false);
            var text = resultDoc
                .MainDocumentPart!.GetXDocument()
                .Descendants(w + "t")
                .Select(t => (string)t)
                .StringConcatenate();
            await Assert.That(text).Contains("Hello, Alice!");
        }
        finally
        {
            DocumentAssembler.UnregisterCustomHandler(elementName);
        }
    }

    /// <summary>DA501 — a custom handler that throws maps to a template error.</summary>
    [Test]
    public async Task DA501_CustomHandler_ExceptionSetsTemplateError()
    {
        const string elementName = "DA501_BrokenHandler";
        try
        {
            DocumentAssembler.RegisterCustomHandler(
                elementName,
                schemaXsd: null,
                handler: (_, _, _) => throw new InvalidOperationException("Simulated handler failure")
            );

            var directive = $"<#<{elementName}/>#>";
            var docxBytes = BuildMinimalDocxWithDirective(directive);
            var wmlTemplate = new WmlDocument("test501.docx", docxBytes);
            var xmlData = XElement.Parse("<Root/>");

            var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

            await Assert.That(hasError).IsTrue();

            using var resultStream = new MemoryStream(result.DocumentByteArray);
            using var resultDoc = WordprocessingDocument.Open(resultStream, false);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var allText = resultDoc
                .MainDocumentPart!.GetXDocument()
                .Descendants(w + "t")
                .Select(t => (string)t)
                .StringConcatenate();
            await Assert.That(allText).Contains("Simulated handler failure");
        }
        finally
        {
            DocumentAssembler.UnregisterCustomHandler(elementName);
        }
    }

    /// <summary>DA502 — a handler returning null silently removes the element.</summary>
    [Test]
    public async Task DA502_CustomHandler_NullReturnRemovesElement()
    {
        const string elementName = "DA502_NullHandler";
        try
        {
            DocumentAssembler.RegisterCustomHandler(elementName, schemaXsd: null, handler: (_, _, _) => null);

            var directive = $"<#<{elementName}/>#>";
            var docxBytes = BuildMinimalDocxWithDirective(directive);
            var wmlTemplate = new WmlDocument("test502.docx", docxBytes);
            var xmlData = XElement.Parse("<Root/>");

            var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

            await Assert.That(hasError).IsFalse();

            using var resultStream = new MemoryStream(result.DocumentByteArray);
            using var resultDoc = WordprocessingDocument.Open(resultStream, false);
            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            var allText = resultDoc
                .MainDocumentPart!.GetXDocument()
                .Descendants(w + "t")
                .Select(t => (string)t)
                .StringConcatenate()
                .Trim();
            await Assert.That(allText).IsEqualTo(string.Empty);
        }
        finally
        {
            DocumentAssembler.UnregisterCustomHandler(elementName);
        }
    }

    /// <summary>DA503 — after UnregisterCustomHandler the element is treated as invalid XML.</summary>
    [Test]
    public async Task DA503_CustomHandler_AfterUnregisterTreatedAsError()
    {
        const string elementName = "DA503_TempHandler";

        DocumentAssembler.RegisterCustomHandler(elementName, schemaXsd: null, handler: (_, _, _) => null);
        DocumentAssembler.UnregisterCustomHandler(elementName);

        var directive = $"<#<{elementName}/>#>";
        var docxBytes = BuildMinimalDocxWithDirective(directive);
        var wmlTemplate = new WmlDocument("test503.docx", docxBytes);
        var xmlData = XElement.Parse("<Root/>");

        var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);

        await Assert.That(hasError).IsTrue();
    }

    /// <summary>DA504 — a custom handler with a schema validates directive attributes.</summary>
    [Test]
    public async Task DA504_CustomHandler_SchemaValidationRejectsInvalidAttributes()
    {
        const string elementName = "DA504_SchemaHandler";
        const string schema =
            @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified'
                xmlns:xs='http://www.w3.org/2001/XMLSchema'>
              <xs:element name='DA504_SchemaHandler'>
                <xs:complexType>
                  <xs:attribute name='Select' type='xs:string' use='required' />
                </xs:complexType>
              </xs:element>
            </xs:schema>";
        try
        {
            DocumentAssembler.RegisterCustomHandler(
                elementName,
                schemaXsd: schema,
                handler: (directive, data, part) =>
                {
                    var value = (string)data.XPathSelectElement((string)directive.Attribute("Select"));
                    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    return new XElement(w + "p", new XElement(w + "r", new XElement(w + "t", value ?? "")));
                }
            );

            // Valid: has required Select attribute
            var validDirective = $"<#<{elementName} Select=\"./Name\"/>#>";
            var docxBytes = BuildMinimalDocxWithDirective(validDirective);
            var wmlTemplate = new WmlDocument("test504.docx", docxBytes);
            var xmlData = XElement.Parse("<Root><Name>Schema Test</Name></Root>");

            var result = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var hasError);
            await Assert.That(hasError).IsFalse();

            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            using var resultStream = new MemoryStream(result.DocumentByteArray);
            using var resultDoc = WordprocessingDocument.Open(resultStream, false);
            var text = resultDoc
                .MainDocumentPart!.GetXDocument()
                .Descendants(w + "t")
                .Select(t => (string)t)
                .StringConcatenate();
            await Assert.That(text).Contains("Schema Test");
        }
        finally
        {
            DocumentAssembler.UnregisterCustomHandler(elementName);
        }
    }
}
