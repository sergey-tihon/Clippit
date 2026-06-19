using System.Text;
using System.Xml.Linq;
using Clippit.Cli.Infrastructure;
using Clippit.Word;
using SkiaSharp;

namespace Clippit.Cli.Commands.Word.ToHtml;

internal static class WordToHtmlService
{
    public static ConvertResult Execute(
        InputSource input,
        OutputTarget output,
        string? pageTitle,
        string? additionalCss,
        string cssPrefix,
        bool fabricateCss,
        bool inlineImages
    )
    {
        // Read source document
        byte[] docBytes;
        using (var stream = input.OpenSeekable())
        using (var memoryStream = new MemoryStream())
        {
            stream.CopyTo(memoryStream);
            docBytes = memoryStream.ToArray();
        }

        var wmlDoc = new WmlDocument(input.DisplayName, docBytes);

        // Determine image directory name for external files mode
        string? imageDirectoryName = null;
        if (!inlineImages && !output.IsStdout)
        {
            var htmlFile = new FileInfo(output.DisplayPath);
            var dirName = htmlFile.Name[..^htmlFile.Extension.Length] + "_files";
            imageDirectoryName = htmlFile.DirectoryName is not null
                ? Path.Combine(htmlFile.DirectoryName, dirName)
                : Path.Combine(Directory.GetCurrentDirectory(), dirName);
        }

        var imageCounter = 0;

        // Build converter settings
        var settings = new WmlToHtmlConverterSettings
        {
            PageTitle = pageTitle ?? input.LogicalName,
            AdditionalCss = additionalCss ?? string.Empty,
            CssClassPrefix = cssPrefix,
            FabricateCssClasses = fabricateCss,
            RestrictToSupportedLanguages = false,
            RestrictToSupportedNumberingFormats = false,
            ImageHandler = imageInfo =>
            {
                imageCounter++;

                if (inlineImages)
                {
                    return CreateInlineImage(imageInfo);
                }

                if (output.IsStdout)
                {
                    throw CliException.InvalidArguments(
                        "--inline-images is required when writing HTML to stdout for documents with images."
                    );
                }

                // Write image to external files directory
                return ImageHelper.DefaultImageHandler(imageInfo, imageDirectoryName!, imageCounter);
            },
        };

        // Convert
        var htmlElement = WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);

        // Produce HTML document with <!DOCTYPE html>
        var html = new XDocument(new XDocumentType("html", null, null, null), htmlElement);
        var htmlString = html.ToString(SaveOptions.DisableFormatting);

        var htmlBytes = Encoding.UTF8.GetBytes(htmlString);

        // Write output
        string? tempPath = null;
        Stream outputStream;
        try
        {
            output.EnsureCanWrite(force: true, "output");
            output.EnsureDirectoryExists();
            outputStream = output.OpenWrite(out tempPath);
        }
        catch (Exception ex) when (ex is not CliException)
        {
            throw CliException.OutputError($"Could not open output for writing: {ex.Message}");
        }

        try
        {
            outputStream.Write(htmlBytes, 0, htmlBytes.Length);
            outputStream.Flush();

            if (output.IsStdout)
            {
                output.Flush(outputStream);
            }
        }
        finally
        {
            outputStream.Dispose();
        }

        try
        {
            if (!output.IsStdout)
            {
                output.Commit(tempPath);
            }
        }
        finally
        {
            OutputTarget.DeleteTemp(tempPath);
        }

        return new ConvertResult
        {
            Input = input.DisplayName,
            Output = output.DisplayPath,
            OutputSize = htmlBytes.Length,
        };
    }

    private static XElement CreateInlineImage(ImageInfo imageInfo)
    {
        var extension = imageInfo.ContentType.Split('/')[1];
        var imageEncoder = ImageHelper.GetEncoder(extension, out extension);

        if (imageEncoder is null)
            return null!;

        string base64;
        try
        {
            using var ms = new MemoryStream();
            using var image = SKImage.FromBitmap(imageInfo.Image);
            using var data = image.Encode(imageEncoder.Value, quality: 80);
            data.SaveTo(ms);
            base64 = Convert.ToBase64String(ms.ToArray());
        }
        catch (System.Runtime.InteropServices.ExternalException)
        {
            return null!;
        }

        var mimeType = "image/" + extension;
        var imageSource = $"data:{mimeType};base64,{base64}";
        var img = new XElement(
            Xhtml.img,
            new XAttribute(NoNamespace.src, imageSource),
            imageInfo.ImgStyleAttribute,
            imageInfo.AltText is not null ? new XAttribute(NoNamespace.alt, imageInfo.AltText) : null
        );
        return img;
    }
}
