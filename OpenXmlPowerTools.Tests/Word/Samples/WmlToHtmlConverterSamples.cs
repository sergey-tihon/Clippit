using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Word.Samples
{
    public class WmlToHtmlConverterSamples : TestsBase
    {
        public WmlToHtmlConverterSamples(ITestOutputHelper log) : base(log)
        {
        }
        
        private static string RootFolder => "../../../Word/Samples/WmlToHtmlConverter/";

        [Fact]
        public void Sample1()
        {
            foreach (var file in Directory.GetFiles(RootFolder, "*.docx"))
            {
                ConvertToHtmlWithExternalFiles(file, TempDir);
            }
        }

        private void ConvertToHtmlWithExternalFiles(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
            Log.WriteLine(fi.Name);
            
            using var memoryStream = new MemoryStream();
            var byteArray = File.ReadAllBytes(fi.FullName);
            memoryStream.Write(byteArray, 0, byteArray.Length);

            using var wDoc = WordprocessingDocument.Open(memoryStream, true);
            var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
            if (!string.IsNullOrEmpty(outputDirectory))
            {
                var di = new DirectoryInfo(outputDirectory);
                if (!di.Exists)
                {
                    throw new OpenXmlPowerToolsException("Output directory does not exist");
                }
                destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
            }
            var imageDirectoryName = destFileName.FullName.Substring(0, destFileName.FullName.Length - 5) + "_files";
            var imageCounter = 0;

            var pageTitle = fi.FullName;
            var part = wDoc.CoreFilePropertiesPart;
            if (part is not null)
            {
                pageTitle = (string)part.GetXDocument().Descendants(DC.title).FirstOrDefault() ?? fi.FullName;
            }

            // TODO: Determine max-width from size of content area.
            var settings = new WmlToHtmlConverterSettings
            {
                AdditionalCss = "body { margin: 1cm auto; max-width: 20cm; padding: 0; }",
                PageTitle = pageTitle,
                FabricateCssClasses = true,
                CssClassPrefix = "pt-",
                RestrictToSupportedLanguages = false,
                RestrictToSupportedNumberingFormats = false,
                ImageHandler = imageInfo =>
                {
                    var localDirInfo = new DirectoryInfo(imageDirectoryName);
                    if (!localDirInfo.Exists)
                        localDirInfo.Create();
                    ++imageCounter;
                    var extension = imageInfo.ContentType.Split('/')[1].ToLower();
                    ImageFormat imageFormat = null;
                    if (extension == "png")
                        imageFormat = ImageFormat.Png;
                    else if (extension == "gif")
                        imageFormat = ImageFormat.Gif;
                    else if (extension == "bmp")
                        imageFormat = ImageFormat.Bmp;
                    else if (extension == "jpeg")
                        imageFormat = ImageFormat.Jpeg;
                    else if (extension == "tiff")
                    {
                        // Convert tiff to gif.
                        extension = "gif";
                        imageFormat = ImageFormat.Gif;
                    }
                    else if (extension == "x-wmf")
                    {
                        extension = "wmf";
                        imageFormat = ImageFormat.Wmf;
                    }

                    // If the image format isn't one that we expect, ignore it,
                    // and don't return markup for the link.
                    if (imageFormat == null)
                        return null;

                    var imageFileName = imageDirectoryName + "/image" +
                                        imageCounter.ToString() + "." + extension;
                    try
                    {
                        imageInfo.Bitmap.Save(imageFileName, imageFormat);
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        return null;
                    }
                    var imageSource = localDirInfo.Name + "/image" +
                                      imageCounter + "." + extension;

                    var img = new XElement(Xhtml.img,
                        new XAttribute(NoNamespace.src, imageSource),
                        imageInfo.ImgStyleAttribute,
                        imageInfo.AltText != null ?
                            new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                    return img;
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
        }
        
        
        [Fact]
        public void Sample2()
        {
            foreach (var file in Directory.GetFiles(RootFolder, "*.docx"))
            {
                ConvertToHtmlWithEmbeddedImages(file, TempDir);
            }
        }

        private void ConvertToHtmlWithEmbeddedImages(string file, string outputDirectory)
        {
            var fi = new FileInfo(file);
            Log.WriteLine(fi.Name);
            
            using var memoryStream = new MemoryStream();
            var byteArray = File.ReadAllBytes(fi.FullName);
            memoryStream.Write(byteArray, 0, byteArray.Length);
            
            using var wDoc = WordprocessingDocument.Open(memoryStream, true);
            var destFileName = new FileInfo(fi.Name.Replace(".docx", ".html"));
            if (!string.IsNullOrEmpty(outputDirectory))
            {
                var di = new DirectoryInfo(outputDirectory);
                if (!di.Exists)
                {
                    throw new OpenXmlPowerToolsException("Output directory does not exist");
                }

                destFileName = new FileInfo(Path.Combine(di.FullName, destFileName.Name));
            }

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
                    var extension = imageInfo.ContentType.Split('/')[1].ToLower();
                    ImageFormat imageFormat = null;
                    if (extension == "png")
                        imageFormat = ImageFormat.Png;
                    else if (extension == "gif")
                        imageFormat = ImageFormat.Gif;
                    else if (extension == "bmp")
                        imageFormat = ImageFormat.Bmp;
                    else if (extension == "jpeg")
                        imageFormat = ImageFormat.Jpeg;
                    else if (extension == "tiff")
                    {
                        // Convert tiff to gif.
                        extension = "gif";
                        imageFormat = ImageFormat.Gif;
                    }
                    else if (extension == "x-wmf")
                    {
                        extension = "wmf";
                        imageFormat = ImageFormat.Wmf;
                    }

                    // If the image format isn't one that we expect, ignore it,
                    // and don't return markup for the link.
                    if (imageFormat == null)
                        return null;

                    string base64 = null;
                    try
                    {
                        using var ms = new MemoryStream();
                        imageInfo.Bitmap.Save(ms, imageFormat);
                        var ba = ms.ToArray();
                        base64 = System.Convert.ToBase64String(ba);
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        return null;
                    }

                    var format = imageInfo.Bitmap.RawFormat;
                    var codec = ImageCodecInfo.GetImageDecoders().First(c => c.FormatID == format.Guid);
                    var mimeType = codec.MimeType;

                    var imageSource = $"data:{mimeType};base64,{base64}";

                    var img = new XElement(Xhtml.img,
                        new XAttribute(NoNamespace.src, imageSource),
                        imageInfo.ImgStyleAttribute,
                        imageInfo.AltText != null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null);
                    return img;
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
        }
    }
}
