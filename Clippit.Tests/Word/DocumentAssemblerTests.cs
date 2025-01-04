// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml;
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace Clippit.Tests.Word
{
    public class DocumentAssemblerTests : TestsBase
    {
        public DocumentAssemblerTests(ITestOutputHelper log)
            : base(log)
        {
            _sourceDir = new DirectoryInfo("../../../../TestFiles/DA/");
        }

        private readonly DirectoryInfo _sourceDir;

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
        [InlineData("DA289A-xhtml-formatting.docx", "DA-html-input.xml", false)]
        [InlineData("DA289B-html-not-supported.docx", "DA-html-input.xml", true)]
        [InlineData("DA289C-not-well-formed-xhtml.docx", "DA-html-input.xml", true)]
        public void DA101(string name, string data, bool err)
        {
            var templateDocx = new FileInfo(Path.Combine(_sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(_sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmlData = XElement.Load(dataFile.FullName);

            var afterAssembling = DocumentAssembler.AssembleDocument(
                wmlTemplate,
                xmlData,
                out var returnedTemplateError
            );
            var assembledDocx = new FileInfo(
                Path.Combine(TempDir, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx"))
            );
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
        [InlineData("DA289-xhtml-formatting.docx", "DA289-invalid.xml", 0, true)]
        [InlineData("DA289-xhtml-formatting.docx", "DA289-not-well-formed.xml", 0, true)]
        public void DA289(string name, string data, int parasInContent, bool err)
        {
            var templateDocx = new FileInfo(Path.Combine(_sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(_sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmlData = XElement.Load(dataFile.FullName);

            var wmlResult = DocumentAssembler.AssembleDocument(wmlTemplate, xmlData, out var returnedTemplateError);
            var assembledDocx = new FileInfo(
                Path.Combine(TempDir, data.Replace(".xml", "-processed-by-DocumentAssembler.docx"))
            );
            wmlResult.SaveAs(assembledDocx.FullName);

            Validate(assembledDocx);
            Assert.Equal(err, returnedTemplateError);

            // if we are not expecting an error then verify that we have the same number of paragraphs and that
            // the paragraph properties from source and target are the same
            if (!err)
            {
                IList<XElement> sourceParas = wmlTemplate.MainDocumentPart.Element(W.body).Descendants(W.p).ToList();
                IList<XElement> targetParas = wmlResult.MainDocumentPart.Element(W.body).Descendants(W.p).ToList();

                // Check we have the expected number of paragraphs
                // Expected document structure is:
                //   Heading paragraph (1 line)
                //   Empty paragraph (1 line)
                //   Escaped HTML paragraph (potential multi-line)
                //   CDATA paragraph (potential multi-line)

                int expectedParas = sourceParas.Count + (2 * parasInContent) - 2;
                Assert.Equal(expectedParas, targetParas.Count);

                var equalityComparer = new XNodeEqualityComparer();
                int paraOffset = 0;

                for (var i = 0; i < sourceParas.Count(); i++)
                {
                    var parasToCompare = i <= 1 ? 1 : parasInContent;
                    var sourceProps = sourceParas[i].Element(W.pPr);

                    for (var j = i + paraOffset; j < i + paraOffset + parasToCompare; j++)
                    {
                        var targetProps = targetParas[j].Element(W.pPr);
                        if (sourceProps == null && targetProps == null)
                        {
                            continue;
                        }

                        Assert.True(equalityComparer.Equals(sourceProps, targetProps));
                    }

                    // update paragraph offset versus source when we have processed multi-line content
                    if (parasToCompare > 1)
                    {
                        paraOffset += parasToCompare - 1;
                    }
                }
            }
        }

        [Theory]
        [InlineData("DA259-MultiLineContents.docx", "DA-Data.xml", false)]
        public void DA259(string name, string data, bool err)
        {
            DA101(name, data, err);
            var assembledDocx = new FileInfo(
                Path.Combine(TempDir, name.Replace(".docx", "-processed-by-DocumentAssembler.docx"))
            );
            var afterAssembling = new WmlDocument(assembledDocx.FullName);
            var brCount = afterAssembling.MainDocumentPart.Element(W.body).Elements(W.p).Count();
            Assert.Equal(6, brCount);
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
            Assert.Throws<OpenXmlPowerToolsException>(() =>
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
            DA101(name, data, err);
            var assembledDocx = new FileInfo(
                Path.Combine(TempDir, name.Replace(".docx", "-processed-by-DocumentAssembler.docx"))
            );
            var afterAssembling = new WmlDocument(assembledDocx.FullName);

            var descendants = afterAssembling.MainDocumentPart.Value;

            Assert.False(descendants.Contains(">"), "Found > on text");
        }

        [Theory]
        [InlineData("DA-xmlerror.docx", "DA-xmlerror.xml")]
        public void DAXmlError(string name, string data)
        {
            var templateDocx = new FileInfo(Path.Combine(_sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(_sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmlData = XElement.Load(dataFile.FullName);

            var afterAssembling = DocumentAssembler.AssembleDocument(
                wmlTemplate,
                xmlData,
                out var returnedTemplateError
            );
            var assembledDocx = new FileInfo(
                Path.Combine(TempDir, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx"))
            );
            afterAssembling.SaveAs(assembledDocx.FullName);
        }

        [Theory]
        [InlineData("DA025-TemplateDocument.docx", "DA-Data.xml", false)]
        public void DA103_UseXmlDocument(string name, string data, bool err)
        {
            var templateDocx = new FileInfo(Path.Combine(_sourceDir.FullName, name));
            var dataFile = new FileInfo(Path.Combine(_sourceDir.FullName, data));

            var wmlTemplate = new WmlDocument(templateDocx.FullName);
            var xmlData = new XmlDocument();
            xmlData.Load(dataFile.FullName);

            var afterAssembling = DocumentAssembler.AssembleDocument(
                wmlTemplate,
                xmlData,
                out var returnedTemplateError
            );
            var assembledDocx = new FileInfo(
                Path.Combine(TempDir, templateDocx.Name.Replace(".docx", "-processed-by-DocumentAssembler.docx"))
            );
            afterAssembling.SaveAs(assembledDocx.FullName);

            Validate(assembledDocx);
            Assert.Equal(err, returnedTemplateError);
        }

        private void Validate(FileInfo fi)
        {
            using var wDoc = WordprocessingDocument.Open(fi.FullName, false);
            Validate(wDoc, s_expectedErrors);
        }

        private static readonly List<string> s_expectedErrors =
            new()
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
            };
    }
}

#endif
