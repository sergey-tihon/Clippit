using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Html.Samples
{
    public class HtmlToWmlConverterSamples : TestsBase
    {
        public HtmlToWmlConverterSamples(ITestOutputHelper log) : base(log)
        {
        }

        private static string GetFilePath(string path) =>
            Path.Combine("../../../Html/Samples/HtmlToWmlConverter/", path);

        [Fact]
        public void Sample1()
        {
            var files = Directory.GetFiles(GetFilePath("Sample1"), "*.html");
            foreach (var file in files)
            {
                ConvertToDocx(file, TempDir);
            }
        }

        [Fact]
        public void Sample2()
        {
            var templateDoc = new FileInfo(GetFilePath("Sample2/TemplateDocument.docx"));
            var dataFile = new FileInfo(Path.Combine(TempDir, "Data.xml"));

            // The following method generates a large data file with random data.
            // In a real world scenario, this is where you would query your data source and produce XML that will drive your document generation process.
            var numberOfDocumentsToGenerate = 100;
            XElement data = GenerateDataFromDataSource(dataFile, numberOfDocumentsToGenerate);

            var wmlDoc = new WmlDocument(templateDoc.FullName);
            var count = 1;
            foreach (var customer in data.Elements("Customer"))
            {
                var assembledDoc = new FileInfo(Path.Combine(TempDir, $"Letter-{count++:0000}.docx"));
                Log.WriteLine("Generating {0}", assembledDoc.Name);
                var wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, customer, out var templateError);
                if (templateError)
                {
                    Log.WriteLine("Errors in template.");
                    Log.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
                }
                wmlAssembledDoc.SaveAs(assembledDoc.FullName);

                Log.WriteLine("Converting to HTML {0}", assembledDoc.Name);
                var htmlFileName = ConvertToHtml(assembledDoc.FullName, TempDir);

                Log.WriteLine("Converting back to DOCX {0}", htmlFileName.Name);
                ConvertToDocx(htmlFileName.FullName, TempDir);
            }
        }

        private void ConvertToDocx(string file, string destinationDir)
        {
            var s_ProduceAnnotatedHtml = true;

            var sourceHtmlFi = new FileInfo(file);
            Log.WriteLine("Converting " + sourceHtmlFi.Name);
            var sourceImageDi = new DirectoryInfo(destinationDir);

            var destCssFi = new FileInfo(Path.Combine(destinationDir, sourceHtmlFi.Name.Replace(".html", "-2.css")));
            var destDocxFi = new FileInfo(Path.Combine(destinationDir,
                sourceHtmlFi.Name.Replace(".html", "-3-ConvertedByHtmlToWml.docx")));
            var annotatedHtmlFi =
                new FileInfo(Path.Combine(destinationDir, sourceHtmlFi.Name.Replace(".html", "-4-Annotated.txt")));

            var html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceHtmlFi);

            var usedAuthorCss =
                HtmlToWmlConverter.CleanUpCss((string)html.Descendants()
                    .FirstOrDefault(d => d.Name.LocalName.ToLower() == "style"));
            File.WriteAllText(destCssFi.FullName, usedAuthorCss);

            var settings = HtmlToWmlConverter.GetDefaultSettings();
            // image references in HTML files contain the path to the subdir that contains the images, so base URI is the name of the directory
            // that contains the HTML files
            settings.BaseUriForImages = sourceHtmlFi.DirectoryName;

            var doc = HtmlToWmlConverter.ConvertHtmlToWml(defaultCss, usedAuthorCss, userCss, html, settings, null,
                s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null);
            doc.SaveAs(destDocxFi.FullName);
        }

        static string defaultCss =
            @"html, address,
blockquote,
body, dd, div,
dl, dt, fieldset, form,
frame, frameset,
h1, h2, h3, h4,
h5, h6, noframes,
ol, p, ul, center,
dir, hr, menu, pre { display: block; unicode-bidi: embed }
li { display: list-item }
head { display: none }
table { display: table }
tr { display: table-row }
thead { display: table-header-group }
tbody { display: table-row-group }
tfoot { display: table-footer-group }
col { display: table-column }
colgroup { display: table-column-group }
td, th { display: table-cell }
caption { display: table-caption }
th { font-weight: bolder; text-align: center }
caption { text-align: center }
body { margin: auto; }
h1 { font-size: 2em; margin: auto; }
h2 { font-size: 1.5em; margin: auto; }
h3 { font-size: 1.17em; margin: auto; }
h4, p,
blockquote, ul,
fieldset, form,
ol, dl, dir,
menu { margin: auto }
a { color: blue; }
h5 { font-size: .83em; margin: auto }
h6 { font-size: .75em; margin: auto }
h1, h2, h3, h4,
h5, h6, b,
strong { font-weight: bolder }
blockquote { margin-left: 40px; margin-right: 40px }
i, cite, em,
var, address { font-style: italic }
pre, tt, code,
kbd, samp { font-family: monospace }
pre { white-space: pre }
button, textarea,
input, select { display: inline-block }
big { font-size: 1.17em }
small, sub, sup { font-size: .83em }
sub { vertical-align: sub }
sup { vertical-align: super }
table { border-spacing: 2px; }
thead, tbody,
tfoot { vertical-align: middle }
td, th, tr { vertical-align: inherit }
s, strike, del { text-decoration: line-through }
hr { border: 1px inset }
ol, ul, dir,
menu, dd { margin-left: 40px }
ol { list-style-type: decimal }
ol ul, ul ol,
ul ul, ol ol { margin-top: 0; margin-bottom: 0 }
u, ins { text-decoration: underline }
br:before { content: ""\A""; white-space: pre-line }
center { text-align: center }
:link, :visited { text-decoration: underline }
:focus { outline: thin dotted invert }
/* Begin bidirectionality settings (do not change) */
BDO[DIR=""ltr""] { direction: ltr; unicode-bidi: bidi-override }
BDO[DIR=""rtl""] { direction: rtl; unicode-bidi: bidi-override }
*[DIR=""ltr""] { direction: ltr; unicode-bidi: embed }
*[DIR=""rtl""] { direction: rtl; unicode-bidi: embed }

";

        static string userCss = @"";
        
        private static string[] s_productNames = {
            "Unicycle",
            "Bicycle",
            "Tricycle",
            "Skateboard",
            "Roller Blades",
            "Hang Glider",
        };

        private static XElement GenerateDataFromDataSource(FileInfo dataFi, int numberOfDocumentsToGenerate)
        {
            var customers = new XElement("Customers");
            var r = new Random();
            for (var i = 0; i < numberOfDocumentsToGenerate; ++i)
            {
                var customer = new XElement("Customer",
                    new XElement("CustomerID", i + 1),
                    new XElement("Name", "Eric White"),
                    new XElement("HighValueCustomer", r.Next(2) == 0 ? "True" : "False"),
                    new XElement("Orders"));
                var orders = customer.Element("Orders");
                var numberOfOrders = r.Next(10) + 1;
                for (var j = 0; j < numberOfOrders; j++)
                {
                    var order = new XElement("Order",
                        new XAttribute("Number", j + 1),
                        new XElement("ProductDescription", s_productNames[r.Next(s_productNames.Length)]),
                        new XElement("Quantity", r.Next(10)),
                        new XElement("OrderDate", "September 26, 2015"));
                    orders.Add(order);
                }
                customers.Add(customer);
            }
            customers.Save(dataFi.FullName);
            return customers;
        }
        
        public static FileInfo ConvertToHtml(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
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
            var settings = new WmlToHtmlConverterSettings()
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
                }
            };
            var htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);

            // Produce HTML document with <!DOCTYPE html > declaration to tell the browser
            // we are using HTML5.
            var html = new XDocument(
                new XDocumentType("html", null, null, null),
                htmlElement);

            // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
            // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
            // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
            // for detailed explanation.
            //
            // If you further transform the XML tree returned by ConvertToHtmlTransform, you
            // must do it correctly, or entities will not be serialized properly.

            var htmlString = html.ToString(SaveOptions.DisableFormatting);
            File.WriteAllText(destFileName.FullName, htmlString, Encoding.UTF8);

            return destFileName;
        }
    }
}
