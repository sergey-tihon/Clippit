using System.Xml;
using System.Xml.Linq;
using Clippit.Cli.Infrastructure;
using Clippit.Html;

namespace Clippit.Cli.Commands.Word.FromHtml;

internal static class WordFromHtmlService
{
    public static WordConvertResult Execute(
        InputSource input,
        OutputTarget output,
        string? cssFilePath,
        string? defaultCssFilePath,
        string? userCss,
        string? baseUri,
        string? majorFont,
        string? minorFont,
        double? fontSize
    )
    {
        // Read HTML content
        string htmlContent;
        using (var stream = input.OpenSeekable())
        using (var reader = new StreamReader(stream))
        {
            htmlContent = reader.ReadToEnd();
        }

        // Parse HTML as XElement (use XDocument first so <!DOCTYPE html> is accepted)
        XElement xhtml;
        try
        {
            xhtml = XDocument.Parse(htmlContent, LoadOptions.PreserveWhitespace).Root!;
        }
        catch (XmlException ex)
        {
            throw CliException.InvalidFormat(
                $"The HTML file is not well-formed XML and could not be parsed: {ex.Message}. "
                    + "The HtmlToWmlConverter requires strict XML-compatible HTML (all tags closed, attributes quoted)."
            );
        }

        // Extract or load author CSS
        string authorCss;
        if (cssFilePath is not null)
        {
            try
            {
                authorCss = HtmlToWmlConverter.CleanUpCss(File.ReadAllText(cssFilePath));
            }
            catch (Exception ex) when (ex is not CliException)
            {
                throw CliException.FileNotFound($"Author CSS file not found or could not be read: {cssFilePath}");
            }
        }
        else
        {
            // Extract from <style> element in HTML
            var styleElement = xhtml
                .Descendants()
                .FirstOrDefault(d => d.Name.LocalName.Equals("style", StringComparison.OrdinalIgnoreCase));

            authorCss = styleElement is not null ? HtmlToWmlConverter.CleanUpCss(styleElement.Value) : string.Empty;
        }

        // Load default CSS
        string defaultCss;
        if (defaultCssFilePath is not null)
        {
            try
            {
                defaultCss = HtmlToWmlConverter.CleanUpCss(File.ReadAllText(defaultCssFilePath));
            }
            catch (Exception ex) when (ex is not CliException)
            {
                throw CliException.FileNotFound(
                    $"Default CSS file not found or could not be read: {defaultCssFilePath}"
                );
            }
        }
        else
        {
            defaultCss = HtmlToWmlConverter.CleanUpCss(HtmlToWmlConverter.DefaultCss);
        }

        // Clean user CSS
        var cleanedUserCss = HtmlToWmlConverter.CleanUpCss(userCss ?? string.Empty);

        // Build converter settings
        var settings = HtmlToWmlConverter.GetDefaultSettings();

        if (majorFont is not null)
            settings.MajorLatinFont = majorFont;

        if (minorFont is not null)
            settings.MinorLatinFont = minorFont;

        if (fontSize.HasValue)
            settings.DefaultFontSize = fontSize.Value;

        if (baseUri is not null)
            settings.BaseUriForImages = baseUri;

        // Convert
        var wmlDoc = HtmlToWmlConverter.ConvertHtmlToWml(defaultCss, authorCss, cleanedUserCss, xhtml, settings);

        var docxBytes = wmlDoc.DocumentByteArray;

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
            outputStream.Write(docxBytes, 0, docxBytes.Length);
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

        return new WordConvertResult
        {
            Input = input.DisplayName,
            Output = output.DisplayPath,
            OutputSize = docxBytes.Length,
        };
    }
}
