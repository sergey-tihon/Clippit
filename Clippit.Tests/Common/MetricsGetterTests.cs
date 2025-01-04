// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Excel;
using Clippit.PowerPoint;
using Clippit.Word;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace Clippit.Tests.Common
{
    public class MetricsGetterTests : TestsBase
    {
        public MetricsGetterTests(ITestOutputHelper log)
            : base(log) { }

        [Theory]
        [InlineData("Presentation.pptx")]
        [InlineData("Spreadsheet.xlsx")]
        [InlineData("DA/DA001-TemplateDocument.docx")]
        [InlineData("DA/DA002-TemplateDocument.docx")]
        [InlineData("DA/DA003-Select-XPathFindsNoData.docx")]
        [InlineData("DA/DA004-Select-XPathFindsNoDataOptional.docx")]
        [InlineData("DA/DA005-SelectRowData-NoData.docx")]
        [InlineData("DA/DA006-SelectTestValue-NoData.docx")]
        public void MG001(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var fi = new FileInfo(Path.Combine(sourceDir.FullName, name));

            var settings = new MetricsGetterSettings()
            {
                IncludeTextInContentControls = false,
                IncludeXlsxTableCellData = false,
                RetrieveNamespaceList = true,
                RetrieveContentTypeList = true,
            };

            var extension = fi.Extension.ToLower();
            XElement metrics = null;
            if (Util.IsWordprocessingML(extension))
            {
                var wmlDocument = new WmlDocument(fi.FullName);
                metrics = MetricsGetter.GetDocxMetrics(wmlDocument, settings);
            }
            else if (Util.IsSpreadsheetML(extension))
            {
                var smlDocument = new SmlDocument(fi.FullName);
                metrics = MetricsGetter.GetXlsxMetrics(smlDocument, settings);
            }
            else if (Util.IsPresentationML(extension))
            {
                var pmlDocument = new PmlDocument(fi.FullName);
                metrics = MetricsGetter.GetPptxMetrics(pmlDocument, settings);
            }

            Assert.NotNull(metrics);
        }
    }
}

#endif
