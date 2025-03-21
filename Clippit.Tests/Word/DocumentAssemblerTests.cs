// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using Clippit.Word;
using DocumentFormat.OpenXml.Office.Y2022.FeaturePropertyBag;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using static System.Runtime.InteropServices.JavaScript.JSType;

#if !ELIDE_XUNIT_TESTS

namespace Clippit.Tests.Word
{
    public class DocumentAssemblerTests(ITestOutputHelper log) : TestsBase(log)
    {
        private readonly DirectoryInfo _sourceDir = new("../../../../TestFiles/DA/");

        [Theory]
        [InlineData("DA001-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA002-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [InlineData("DA003-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [InlineData("DA004-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [InlineData("DA005-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA006-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA007-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA008-TableElementWithNoTable.docx", "DA-Data.xml", true)]
        [InlineData("DA009-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA010-InvalidXml.docx", "DA-Data.xml", true)]
        [InlineData("DA011-SchemaError.docx", "DA-Data.xml", true)]
        [InlineData("DA012-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [InlineData("DA013-Runs.docx", "DA-Data.xml", false)]
        [InlineData("DA014-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [InlineData("DA015-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [InlineData("DA016-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [InlineData("DA017-FiveRuns.docx", "DA-Data.xml", true)]
        [InlineData("DA018-SmartQuotes.docx", "DA-Data.xml", false)]
        [InlineData("DA019-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [InlineData("DA020-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [InlineData("DA021-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [InlineData("DA022-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA023-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA026-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [InlineData("DA027-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [InlineData("DA028-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [InlineData("DA029-NoDataForCell.docx", "DA-Data.xml", true)]
        [InlineData("DA030-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", false)] // Clippit support multi-value XPath in table cells
        [InlineData("DA031-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [InlineData("DA032-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [InlineData("DA033-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [InlineData("DA034-HeaderFooter.docx", "DA-Data.xml", false)]
        [InlineData("DA035-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA036-SchemaErrorInConditional.docx", "DA-Data.xml", true)]
        [InlineData("DA100-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA101-TemplateDocument.docx", "DA-Data.xml", true)]
        [InlineData("DA102-TemplateDocument.docx", "DA-Data.xml", true)]
        [InlineData("DA201-TemplateDocument.docx", "DA-Data.xml", false)]
        [InlineData("DA202-TemplateDocument.docx", "DA-DataNotHighValueCust.xml", false)]
        [InlineData("DA203-Select-XPathFindsNoData.docx", "DA-Data.xml", true)]
        [InlineData("DA204-Select-XPathFindsNoDataOptional.docx", "DA-Data.xml", false)]
        [InlineData("DA205-SelectRowData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA206-SelectTestValue-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA207-SelectRepeatingData-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA209-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA210-InvalidXml.docx", "DA-Data.xml", true)]
        [InlineData("DA211-SchemaError.docx", "DA-Data.xml", true)]
        [InlineData("DA212-OtherMarkupTypes.docx", "DA-Data.xml", true)]
        [InlineData("DA213-Runs.docx", "DA-Data.xml", false)]
        [InlineData("DA214-TwoRuns-NoValuesSelected.docx", "DA-Data.xml", true)]
        [InlineData("DA215-TwoRunsXmlExceptionInFirst.docx", "DA-Data.xml", true)]
        [InlineData("DA216-TwoRunsSchemaErrorInSecond.docx", "DA-Data.xml", true)]
        [InlineData("DA217-FiveRuns.docx", "DA-Data.xml", true)]
        [InlineData("DA218-SmartQuotes.docx", "DA-Data.xml", false)]
        [InlineData("DA219-RunIsEntireParagraph.docx", "DA-Data.xml", false)]
        [InlineData("DA220-TwoRunsAndNoOtherContent.docx", "DA-Data.xml", true)]
        [InlineData("DA221-NestedRepeat.docx", "DA-DataNestedRepeat.xml", false)]
        [InlineData("DA222-InvalidXPath.docx", "DA-Data.xml", true)]
        [InlineData("DA223-RepeatWOEndRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA226-InvalidRootXmlElement.docx", "DA-Data.xml", true)]
        [InlineData("DA227-XPathErrorInPara.docx", "DA-Data.xml", true)]
        [InlineData("DA228-NoPrototypeRow.docx", "DA-Data.xml", true)]
        [InlineData("DA229-NoDataForCell.docx", "DA-Data.xml", true)]
        [InlineData("DA230-TooMuchDataForCell.docx", "DA-TooMuchDataForCell.xml", false)] // Clippit support multi-value XPath in table cells
        [InlineData("DA231-CellDataInAttributes.docx", "DA-CellDataInAttributes.xml", true)]
        [InlineData("DA232-TooMuchDataForConditional.docx", "DA-TooMuchDataForConditional.xml", true)]
        [InlineData("DA233-ConditionalOnAttribute.docx", "DA-ConditionalOnAttribute.xml", false)]
        [InlineData("DA234-HeaderFooter.docx", "DA-Data.xml", false)]
        [InlineData("DA235-Crashes.docx", "DA-Content-List.xml", false)]
        [InlineData("DA236-Page-Num-in-Footer.docx", "DA-Content-List.xml", false)]
        [InlineData("DA237-SchemaErrorInRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA238-SchemaErrorInConditional.docx", "DA-Data.xml", true)]
        [InlineData("DA239-RunLevelCC-Repeat.docx", "DA-Data.xml", false)]
        [InlineData("DA250-ConditionalWithRichXPath.docx", "DA250-Address.xml", false)]
        [InlineData("DA251-EnhancedTables.docx", "DA-Data.xml", false)]
        [InlineData("DA252-Table-With-Sum.docx", "DA-Data.xml", false)]
        [InlineData("DA253-Table-With-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
        [InlineData("DA254-Table-With-XPath-Sum.docx", "DA-Data.xml", false)]
        [InlineData("DA255-Table-With-XPath-Sum-Run-Level-CC.docx", "DA-Data.xml", false)]
        [InlineData("DA256-NoInvalidDocOnErrorInRun.docx", "DA-Data.xml", true)]
        [InlineData("DA257-OptionalRepeat.docx", "DA-Data.xml", false)]
        [InlineData("DA258-ContentAcceptsCharsAsXPathResult.docx", "DA-Data.xml", false)]
        [InlineData("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        [InlineData("DA260-RunLevelRepeat.docx", "DA-Data.xml", false)]
        [InlineData("DA261-RunLevelConditional.docx", "DA-Data.xml", false)]
        [InlineData("DA262-ConditionalNotMatch.docx", "DA-Data.xml", false)]
        [InlineData("DA263-ConditionalNotMatch.docx", "DA-DataSmallCustomer.xml", false)]
        [InlineData("DA264-InvalidRunLevelRepeat.docx", "DA-Data.xml", true)]
        [InlineData("DA265-RunLevelRepeatWithWhiteSpaceBefore.docx", "DA-Data.xml", false)]
        [InlineData("DA266-RunLevelRepeat-NoData.docx", "DA-Data.xml", true)]
        [InlineData("DA267-Repeat-HorizontalAlignType.docx", "DA-Data.xml", false)]
        [InlineData("DA268-Repeat-VerticalAlignType.docx", "DA-Data.xml", false)]
        [InlineData("DA269-Repeat-InvalidAlignType.docx", "DA-Data.xml", true)]
        [InlineData("DA270-ImageSelect.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA270A-ImageSelect.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA271-ImageSelectWithRepeat.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA271A-ImageSelectWithRepeat.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA272-ImageSelectWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA272A-ImageSelectWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA273-ImageSelectInsideTextBoxWithRepeatVerticalAlign.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA273A-ImageSelectInsideTextBoxWithRepeatVerticalAlign.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA274-ImageSelectInsideTextBoxWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA274A-ImageSelectInsideTextBoxWithRepeatHorizontalAlign.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA275-ImageSelectWithRepeatInvalidAlign.docx", "DA-Data-WithImages.xml", true)]
        [InlineData("DA275A-ImageSelectWithRepeatInvalidAlign.docx", "DA-Data-WithImages.xml", true)]
        [InlineData("DA276-ImageSelectInsideTable.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA276A-ImageSelectInsideTable.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA277-ImageSelectMissingOrInvalidPictureContent.docx", "DA-Data-WithImages.xml", true)]
        [InlineData("DA277A-ImageSelectMissingOrInvalidPictureContent.docx", "DA-Data-WithImages.xml", true)]
        [InlineData("DA278-ImageSelect.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
        [InlineData("DA278A-ImageSelect.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
        [InlineData("DA279-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidMIMEType.xml", true)]
        [InlineData("DA279A-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidMIMEType.xml", true)]
        [InlineData("DA280-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidImageDataFormat.xml", true)]
        [InlineData("DA280A-ImageSelectWithRepeat.docx", "DA-Data-WithImagesInvalidImageDataFormat.xml", true)]
        [InlineData("DA281-ImageSelectExtraWhitespaceBeforeImageContent.docx", "DA-Data-WithImages.xml", true)]
        [InlineData("DA281A-ImageSelectExtraWhitespaceBeforeImageContent.docx", "DA-Data-WithImages.xml", true)]
        [InlineData("DA282-ImageSelectWithHeader.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA282A-ImageSelectWithHeader.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA282-ImageSelectWithHeader.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
        [InlineData("DA282A-ImageSelectWithHeader.docx", "DA-Data-WithImagesInvalidPath.xml", true)]
        [InlineData("DA283-ImageSelectWithFooter.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA283A-ImageSelectWithFooter.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA284-ImageSelectWithHeaderAndFooter.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA284A-ImageSelectWithHeaderAndFooter.docx", "DA-Data-WithImages.xml", false)]
        [InlineData("DA285-ImageSelectNoParagraphFollowedAfterMetadata.docx", "DA-Data-WithImages.xml", true)]
        [InlineData("DA285A-ImageSelectNoParagraphFollowedAfterMetadata.docx", "DA-Data-WithImages.xml", true)]
        [InlineData("DA-I0038-TemplateWithMultipleXPathResults.docx", "DA-I0038-Data.xml", false)]
        public void DA101(string name, string data, bool err)
        {
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);

            Validate(assembledDocx);
            Assert.Equal(err, returnedTemplateError);
        }

        [Theory]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-no-inline-styles.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-bold.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-strong.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-italic.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-emphasis.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-underline.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-bold-underline.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-bold-italic.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-italic-underline.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-subscript.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-superscript.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-strikethrough.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-hyperlink.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-hyperlink-bold.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-hyperlink-italic.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-hyperlink-underline.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-hyperlink-no-protocol.xml", 1, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-multi-paragraph.xml", 3, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-multi-paragraph-with-CRLF.xml", 3, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-multi-paragraph-text-only.xml", 3, false)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-invalid.xml", 0, true)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-not-well-formed.xml", 0, true)]
        public void DA289(string name, string data, int parasInContent, bool err)
        {
            var wmlTemplate = new WmlDocument(Path.Combine(_sourceDir.FullName, name));

            var wmlResult = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(data);
            wmlResult.SaveAs(assembledDocx.FullName);

            Validate(assembledDocx);
            Assert.Equal(err, returnedTemplateError);

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
                Assert.Equal(sourceParas.Count(), targetParas.Count());
                int expectedBreaks = (parasInContent - 1) * 2;
                Assert.Equal(expectedBreaks, targetBreaks.Count());

                var equalityComparer = new XNodeEqualityComparer();
                int paraOffset = 0;

                for (var i = 0; i < sourceParas.Count(); i++)
                {
                    var sourceProps = sourceParas[i].Element(W.pPr);
                    var targetProps = targetParas[i].Element(W.pPr);

                    if (sourceProps == null && targetProps == null)
                    {
                        continue;
                    }

                    Assert.True(equalityComparer.Equals(sourceProps, targetProps));
                }
            }
        }

        [Theory]
        [InlineData("DA290-xhtml-merge-run-formatting.docx", "DA290-xhtml-merge-run-formatting.xml")]
        public void DA290_Merge_Run_Formatting(string name, string data)
        {
            // Act
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);

            // Assert - para count is expected
            List<XElement> paras = afterAssembling.MainDocumentPart.Element(W.body).Descendants(W.p).ToList();
            Assert.Equal(9, paras.Count());

            // Assert - Paragraph 1 Styles
            XElement para = paras[0];
            Assert.Equal("Heading1", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.Equal("16", para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value);
            Assert.Equal("16", para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value);
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("16", x.Element(W.sz).Attribute(W.val).Value);
                    Assert.Equal("16", x.Element(W.szCs).Attribute(W.val).Value);
                }
            );

            // Assert - Paragraph 2 Styles
            para = paras[1];
            Assert.Equal("Heading2", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("Heading2Char", x.Element(W.rStyle).Attribute(W.val).Value);
                }
            );

            // Assert - Paragraph 3 Styles
            para = paras[2];
            Assert.Equal("Heading2CDATA", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("538135", x.Element(W.color).Attribute(W.val).Value);
                }
            );

            // Assert - Paragraph 4 Styles
            para = paras[3];
            Assert.Equal("Heading2CDATA", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.NotNull(para.Element(W.pPr).Element(W.rPr).Element(W.rFonts));
            Assert.Equal("auto", para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value);
            Assert.Equal("22", para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value);
            Assert.Equal("22", para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value);
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("Algerian", x.Element(W.rFonts).Attribute(W.ascii).Value);
                    Assert.NotNull(x.Element(W.i));
                    Assert.NotNull(x.Element(W.iCs));
                    Assert.Equal("single", x.Element(W.u).Attribute(W.val).Value);
                }
            );

            // Assert - Paragraph 5 Styles
            para = paras[4];
            Assert.Equal("Heading2CDATA", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.Equal("C45911", para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value);
            Assert.Equal("14", para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value);
            Assert.Equal("14", para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value);
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("14", x.Element(W.sz).Attribute(W.val).Value);
                    Assert.Equal("14", x.Element(W.szCs).Attribute(W.val).Value);
                }
            );

            // Assert - Paragraph 6 Styles
            para = paras[5];
            Assert.Equal("Heading2CDATA", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.Equal("C45911", para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value);
            Assert.Equal("40", para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value);
            Assert.Equal("40", para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value);
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("40", x.Element(W.sz).Attribute(W.val).Value);
                    Assert.Equal("40", x.Element(W.szCs).Attribute(W.val).Value);
                }
            );

            // Assert - Paragraph 7 Styles
            para = paras[6];
            Assert.Equal("Heading2CDATA", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.Equal("00B0F0", para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value);
            Assert.Equal("40", para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value);
            Assert.Equal("40", para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value);
            Assert.Equal("Algerian", paras[6].Element(W.pPr).Element(W.rPr).Element(W.rFonts).Attribute(W.ascii).Value);
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("Algerian", x.Element(W.rFonts).Attribute(W.ascii).Value);
                    Assert.Equal("00B0F0", x.Element(W.color).Attribute(W.val).Value);
                    Assert.Equal("40", x.Element(W.sz).Attribute(W.val).Value);
                    Assert.Equal("40", x.Element(W.szCs).Attribute(W.val).Value);
                }
            );

            // Assert - Paragraph 8 Styles
            para = paras[7];
            Assert.Equal("Heading2CDATA", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.Equal("C45911", para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value);
            Assert.Equal("32", para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value);
            Assert.Equal("32", para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value);
            Assert.Equal("Algerian", para.Element(W.pPr).Element(W.rPr).Element(W.rFonts).Attribute(W.ascii).Value);
            Assert.Equal("single", para.Element(W.pPr).Element(W.rPr).Element(W.u).Attribute(W.val).Value);
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("Algerian", x.Element(W.rFonts).Attribute(W.ascii).Value);
                    Assert.Equal("single", x.Element(W.u).Attribute(W.val).Value);
                    Assert.Equal("32", x.Element(W.sz).Attribute(W.val).Value);
                    Assert.Equal("32", x.Element(W.szCs).Attribute(W.val).Value);
                }
            );

            // Assert - Paragraph 9 Styles
            para = paras[8];
            Assert.Equal("Heading2CDATA", para.Element(W.pPr).Element(W.pStyle).Attribute(W.val).Value);
            Assert.Equal("538135", para.Element(W.pPr).Element(W.rPr).Element(W.color).Attribute(W.val).Value);
            Assert.Equal("28", para.Element(W.pPr).Element(W.rPr).Element(W.sz).Attribute(W.val).Value);
            Assert.Equal("28", para.Element(W.pPr).Element(W.rPr).Element(W.szCs).Attribute(W.val).Value);
            Assert.Equal("Algerian", para.Element(W.pPr).Element(W.rPr).Element(W.rFonts).Attribute(W.ascii).Value);
            Assert.Equal("single", para.Element(W.pPr).Element(W.rPr).Element(W.u).Attribute(W.val).Value);
            Assert.NotNull(para.Element(W.pPr).Element(W.rPr).Element(W.i));
            Assert.NotNull(para.Element(W.pPr).Element(W.rPr).Element(W.iCs));
            Assert.All(
                para.Descendants(W.r).Elements(W.rPr),
                x =>
                {
                    Assert.Equal("Algerian", x.Element(W.rFonts).Attribute(W.ascii).Value);
                    Assert.Equal("538135", x.Element(W.color).Attribute(W.val).Value);
                    Assert.Equal("single", x.Element(W.u).Attribute(W.val).Value);
                    Assert.Equal("28", x.Element(W.sz).Attribute(W.val).Value);
                    Assert.Equal("28", x.Element(W.szCs).Attribute(W.val).Value);
                    Assert.NotNull(x.Element(W.i));
                    Assert.NotNull(x.Element(W.iCs));
                }
            );
        }

        [Theory]
        [InlineData("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        public void DA259(string name, string data, bool err)
        {
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);

            var brCount = afterAssembling.MainDocumentPart.Element(W.body).Descendants(W.r).Elements(W.br).Count();

            Assert.Equal(4, brCount);
        }

        [Theory]
        [InlineData("DA286-DocumentTemplate-Base-Main.docx", "DA286-DocumentTemplate-Base.xml", false)]
        [InlineData(
            "DA286-DocumentTemplate-MirroredMargins-Main.docx",
            "DA286-DocumentTemplate-MirroredMargins.xml",
            false
        )]
        [InlineData("DA286-DocumentTemplate-NoBreaks-Main.docx", "DA286-DocumentTemplate-NoBreaks.xml", false)]
        [InlineData("DA286-DocumentTemplate-HeaderFooter-Main.docx", "DA286-DocumentTemplate-HeaderFooter.xml", false)]
        [InlineData("DA286-Document-SolarSystem-Main.docx", "DA286-Document-SolarSystem.xml", false)]
        public void DA286(string templateName, string data, bool err)
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

            Validate(assembledDocx);

            Assert.Equal(err, templateError);
        }

        [Theory]
        [InlineData("DA024-TrackedRevisions.docx", "DA-Data.xml")]
        public void DA102_Throws(string name, string data)
        {
            var templateDocx = new FileInfo(Path.Combine(_sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(_sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmldata = XElement.Load(dataFile.FullName);

            WmlDocument afterAssembling;
            _ = Assert.Throws<OpenXmlPowerToolsException>(() =>
            {
                afterAssembling = DocumentAssembler.AssembleDocument(
                    wmlTemplate,
                    xmldata,
                    out var returnedTemplateError
                );
            });
        }

        [Theory]
        [InlineData("DA-TemplateMaior.docx", "DA-templateMaior.xml", false)]
        public void DATemplateMaior(string name, string data, bool err)
        {
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);

            var descendants = afterAssembling.MainDocumentPart.Value;

            Assert.False(descendants.Contains(">"), "Found > on text");
        }

        [Theory]
        [InlineData("DA-xmlerror.docx", "DA-xmlerror.xml")]
        public void DAXmlError(string name, string data)
        {
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);
        }

        [Theory]
        [InlineData("DA025-TemplateDocument.docx", "DA-Data.xml", false)]
        public void DA103_UseXmlDocument(string name, string data, bool err)
        {
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);

            Validate(assembledDocx);
            Assert.Equal(err, returnedTemplateError);
        }

        [Theory]
        [InlineData("DA-Ampersand+LF-Issue.docx", "DA-Ampersand+LF-Issue.xml", false)]
        [InlineData("DA-Ampersand+LF-Issue-With-Controls.docx", "DA-Ampersand+LF-Issue.xml", false)]
        public void DA_Ampersands_And_LineFeeds(string name, string data, bool err)
        {
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);

            // Assert - no errors
            Validate(assembledDocx);
            Assert.Equal(err, returnedTemplateError);

            // Assert - tables is present and correct
            XElement table = afterAssembling.MainDocumentPart.Descendants(W.tbl).SingleOrDefault();
            Assert.NotNull(table);

            // Assert - the second table cell of each table has one paragraph
            IEnumerable<XElement> paras = table.Descendants(W.tc).ElementAt(1).Elements(W.p);
            Assert.True(paras.Count() == 1);

            // Assert - first table paragraph has 2 soft breaks
            Assert.True(paras.ElementAt(0).Elements(W.r).Count() == 5);
            Assert.True(paras.ElementAt(0).Elements(W.r).Elements(W.br).Count() == 2);
        }

        [Theory]
        [InlineData("DA-Tabs-In-Text.docx", "DA-Tabs-In-Text.xml", false)]
        [InlineData("DA-Tabs-In-Text-With-Controls.docx", "DA-Tabs-In-Text.xml", false)]
        public void DA_Tabs_In_Text(string name, string data, bool err)
        {
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);

            // Assert - no errors
            Validate(assembledDocx);
            Assert.Equal(err, returnedTemplateError);

            // Assert - we have four paragraphs
            IEnumerable<XElement> paras = afterAssembling.MainDocumentPart.Descendants(W.p);
            Assert.True(paras.Count() == 4);

            // Assert - first paragraph has 0 tabs
            Assert.True(paras.ElementAt(0).Descendants(W.tab).Count() == 0);

            // Assert - second paragraph has a tab in the first run
            Assert.True(paras.ElementAt(1).Elements(W.r).First().Elements(W.tab).Count() == 1);

            // Assert - third paragraph has a tab in the last run
            Assert.True(paras.ElementAt(2).Elements(W.r).Last().Elements(W.tab).Count() == 1);

            // Assert - fourth paragraph has a tab but not in the first or last run
            Assert.True(paras.ElementAt(3).Descendants(W.tab).Count() == 1);
            Assert.True(paras.ElementAt(3).Elements(W.r).First().Elements(W.tab).Count() == 0);
            Assert.True(paras.ElementAt(3).Elements(W.r).Last().Elements(W.tab).Count() == 0);
        }

        [Theory]
        [InlineData("DA-Issue-95-Template.docx", "DA-Issue-95-Data.xml", false)]
        public void DA_Issue_95_Repro(string name, string data, bool err)
        {
            var afterAssembling = AssembleDocument(name, data, out bool returnedTemplateError);
            FileInfo assembledDocx = GetOutputFile(name);
            afterAssembling.SaveAs(assembledDocx.FullName);

            // Assert - no errors
            Validate(assembledDocx);
            Assert.Equal(err, returnedTemplateError);

            // Assert - tables are present and correct
            IEnumerable<XElement> tables = afterAssembling.MainDocumentPart.Descendants(W.tbl);
            Assert.True(tables.Count() == 4);

            // Assert - the second table cell of each table has one paragraph
            List<XElement> paras = new List<XElement>();
            foreach (XElement table in tables)
            {
                paras.AddRange(table.Descendants(W.tc).ElementAt(1).Elements(W.p));
            }

            Assert.True(paras.Count() == tables.Count());

            // Assert - first tables paragraph has 4 soft breaks
            Assert.True(paras.ElementAt(0).Elements(W.r).Count() == 7);
            Assert.True(paras.ElementAt(0).Elements(W.r).Elements(W.br).Count() == 4);

            // Assert - second tables paragraph has 1 soft breaks
            Assert.True(paras.ElementAt(1).Elements(W.r).Count() == 3);
            Assert.True(paras.ElementAt(1).Elements(W.r).Elements(W.br).Count() == 1);

            // Assert - third tables paragraph has 2 soft breaks
            Assert.True(paras.ElementAt(2).Elements(W.r).Count() == 5);
            Assert.True(paras.ElementAt(2).Elements(W.r).Elements(W.br).Count() == 2);

            // Assert - fourth tables paragraph has 1 soft breaks and two tabs
            Assert.True(paras.ElementAt(3).Elements(W.r).Count() == 5);
            Assert.True(paras.ElementAt(3).Elements(W.r).Elements(W.br).Count() == 1);
            Assert.True(paras.ElementAt(3).Elements(W.r).Elements(W.tab).Count() == 2);
        }

        private void Validate(FileInfo fi)
        {
            using var wDoc = WordprocessingDocument.Open(fi.FullName, false);
            Validate(wDoc, s_expectedErrors);
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
                Path.Combine(
                    TempDir,
                    fileName.Replace(Path.GetExtension(fileName), "-processed-by-DocumentAssembler.docx")
                )
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
        };
    }
}

#endif
