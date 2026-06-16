using System.Text;
using System.Xml.Linq;
using Clippit.Html;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Html.Samples
{
    public class HtmlToWmlConverterSamples() : Clippit.Tests.TestsBase
    {
        private static string GetFilePath(string path) =>
            Path.Combine("../../../Html/Samples/HtmlToWmlConverter/", path);

        [Test]
        public void Sample1()
        {
            var files = Directory.GetFiles(GetFilePath("Sample1"), "*.html");
            foreach (var file in files)
            {
                ConvertToDocx(file, TempDir);
            }
        }

        [Test]
        public void Sample2()
        {
            var templateDoc = new FileInfo(GetFilePath("Sample2/TemplateDocument.docx"));
            var dataFile = new FileInfo(Path.Combine(TempDir, "Data.xml"));
            // The following method generates a large data file with random data.
            // In a real world scenario, this is where you would query your data source and produce XML that will drive your document generation process.
            var numberOfDocumentsToGenerate = 100;
            var data = GenerateDataFromDataSource(dataFile, numberOfDocumentsToGenerate);
            var wmlDoc = new WmlDocument(templateDoc.FullName);
            var count = 1;
            foreach (var customer in data.Elements("Customer"))
            {
                var assembledDoc = new FileInfo(Path.Combine(TempDir, $"Letter-{count++:0000}.docx"));
                Console.WriteLine("Generating {0}", assembledDoc.Name);
                var wmlAssembledDoc = DocumentAssembler.AssembleDocument(wmlDoc, customer, out var templateError);
                if (templateError)
                {
                    Console.WriteLine("Errors in template.");
                    Console.WriteLine("See {0} to determine the errors in the template.", assembledDoc.Name);
                }

                wmlAssembledDoc.SaveAs(assembledDoc.FullName);
                Console.WriteLine("Converting to HTML {0}", assembledDoc.Name);
                var htmlFileName = ConvertToHtml(assembledDoc.FullName, TempDir);
                Console.WriteLine("Converting back to DOCX {0}", htmlFileName.Name);
                ConvertToDocx(htmlFileName.FullName, TempDir);
            }
        }

        private void ConvertToDocx(string file, string destinationDir)
        {
            var s_ProduceAnnotatedHtml = true;
            var sourceHtmlFi = new FileInfo(file);
            Console.WriteLine("Converting " + sourceHtmlFi.Name);
            var sourceImageDi = new DirectoryInfo(destinationDir);
            var destCssFi = new FileInfo(Path.Combine(destinationDir, sourceHtmlFi.Name.Replace(".html", "-2.css")));
            var destDocxFi = new FileInfo(
                Path.Combine(destinationDir, sourceHtmlFi.Name.Replace(".html", "-3-ConvertedByHtmlToWml.docx"))
            );
            var annotatedHtmlFi = new FileInfo(
                Path.Combine(destinationDir, sourceHtmlFi.Name.Replace(".html", "-4-Annotated.txt"))
            );
            var html = HtmlToWmlReadAsXElement.ReadAsXElement(sourceHtmlFi);
            var usedAuthorCss = HtmlToWmlConverter.CleanUpCss(
                (string)html.Descendants().FirstOrDefault(d => d.Name.LocalName.ToLower() == "style")
            );
            File.WriteAllText(destCssFi.FullName, usedAuthorCss);
            var settings = HtmlToWmlConverter.GetDefaultSettings();
            // image references in HTML files contain the path to the subdir that contains the images, so base URI is the name of the directory
            // that contains the HTML files
            settings.BaseUriForImages = sourceHtmlFi.DirectoryName;
            var doc = HtmlToWmlConverter.ConvertHtmlToWml(
                HtmlToWmlConverter.DefaultCss,
                usedAuthorCss,
                userCss,
                html,
                settings,
                null,
                s_ProduceAnnotatedHtml ? annotatedHtmlFi.FullName : null
            );
            doc.SaveAs(destDocxFi.FullName);
        }

        private static readonly string userCss = @"";
        private static readonly string[] s_productNames =
        [
            "Unicycle",
            "Bicycle",
            "Tricycle",
            "Skateboard",
            "Roller Blades",
            "Hang Glider",
        ];

        private static XElement GenerateDataFromDataSource(FileInfo dataFi, int numberOfDocumentsToGenerate)
        {
            var customers = new XElement("Customers");
            var r = new Random();
            for (var i = 0; i < numberOfDocumentsToGenerate; ++i)
            {
                var customer = new XElement(
                    "Customer",
                    new XElement("CustomerID", i + 1),
                    new XElement("Name", "Eric White"),
                    new XElement("HighValueCustomer", r.Next(2) == 0 ? "True" : "False"),
                    new XElement("Orders")
                );
                var orders = customer.Element("Orders");
                var numberOfOrders = r.Next(10) + 1;
                for (var j = 0; j < numberOfOrders; j++)
                {
                    var order = new XElement(
                        "Order",
                        new XAttribute("Number", j + 1),
                        new XElement("ProductDescription", s_productNames[r.Next(s_productNames.Length)]),
                        new XElement("Quantity", r.Next(10)),
                        new XElement("OrderDate", "September 26, 2015")
                    );
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
                },
            };
            var htmlElement = WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
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
            return destFileName;
        }
    }
}
