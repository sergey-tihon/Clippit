using System.Text;
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Clippit.Tests.Html.Samples
{
    /***************************************************************************
    * IMPORTANT NOTE:
    *
    * With versions 4.1 and later, the name of the HtmlConverter class has been
    * changed to WmlToHtmlConverter, to make it orthogonal with HtmlToWmlConverter.
    *
    * There are thin wrapper classes, HtmlConverter, and HtmlConverterSettings,
    * which maintain backwards compat for code that uses the old name.
    *
    * Other than the name change of the classes themselves, the functionality
    * in WmlToHtmlConverter is identical to the old HtmlConverter class.
   ***************************************************************************/

    public class HtmlConverterSamples(ITestOutputHelper log) : TestsBase(log)
    {
        [Theory]
        [InlineData("5DayTourPlanTemplate.docx")]
        [InlineData("Contract.docx")]
        [InlineData("HcTest-01.docx")]
        [InlineData("HcTest-02.docx")]
        [InlineData("HcTest-03.docx")]
        [InlineData("HcTest-04.docx")]
        [InlineData("HcTest-05.docx")]
        [InlineData("HcTest-06.docx")]
        [InlineData("HcTest-07.docx")]
        [InlineData("HcTest-08.docx")]
        [InlineData("Hebrew-01.docx")]
        [InlineData("Hebrew-02.docx")]
        [InlineData("ResumeTemplate.docx")]
        [InlineData("TaskPlanTemplate.docx")]
        /// This example loads each document into a byte array, then into a memory stream, so that the document can be opened for writing without
        /// modifying the source document.
        public void ConvertToHtml(string fileName)
        {
            var fi = new FileInfo(Path.Combine("../../../Html/Samples/HtmlConverter/", fileName));
            var byteArray = File.ReadAllBytes(fi.FullName);
            using var memoryStream = new MemoryStream();
            memoryStream.Write(byteArray, 0, byteArray.Length);
            using var wDoc = WordprocessingDocument.Open(memoryStream, true);
            var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
            destFileName = new FileInfo(Path.Combine(TempDir, destFileName.Name));

            var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
            var imageCounter = 0;

            var pageTitle = fi.FullName;
            var part = wDoc.CoreFilePropertiesPart;
            if (part != null)
            {
                pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
            }

            // TODO: Determine max-width from size of content area.
            var settings = new HtmlConverterSettings()
            {
                AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                PageTitle = pageTitle,
                FabricateCssClasses = true,
                CssClassPrefix = "pt-",
                RestrictToSupportedLanguages = false,
                RestrictToSupportedNumberingFormats = false,
                ImageHandler = imageInfo =>
                {
                    ++imageCounter;
                    return ImageHelper.DefaultImageHandler(imageInfo, imageDirectoryName, imageCounter);
                },
            };
            var htmlElement = HtmlConverter.ConvertToHtml(wDoc, settings);

            // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
            // we are using HTML5.
            var html = new XDocument(new XDocumentType("html", null, null, null), htmlElement);

            // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
            // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
            // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
            // for detailed explanation.
            //
            // If you further transform the XML tree returned by ConvertToHtmlTransform, you
            // must do it correctly, or entities will not be serialized properly.

            var htmlString = html.ToString(SaveOptions.DisableFormatting);
            File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);
        }
    }
}
