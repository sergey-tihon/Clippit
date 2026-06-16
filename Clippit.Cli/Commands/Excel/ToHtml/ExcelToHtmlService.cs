using System.Text;
using System.Xml.Linq;
using Clippit.Cli.Infrastructure;
using Clippit.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Cli.Commands.Excel.ToHtml;

internal static class ExcelToHtmlService
{
    public static ConvertResult Execute(
        InputSource input,
        OutputTarget output,
        string? sheetName,
        string? range,
        string? tableName,
        string? pageTitle,
        string? additionalCss,
        string cssPrefix,
        bool fabricateCss
    )
    {
        // Enforce validation constraints
        if (!string.IsNullOrEmpty(tableName) && (!string.IsNullOrEmpty(sheetName) || !string.IsNullOrEmpty(range)))
        {
            throw CliException.InvalidArguments("--table option cannot be combined with --sheet or --range.");
        }

        if (!string.IsNullOrEmpty(range) && string.IsNullOrEmpty(sheetName))
        {
            throw CliException.InvalidArguments("--range option requires a --sheet to be specified.");
        }

        // Open spreadsheet document
        XElement rangeXml;
        byte[] docBytes;
        using (var stream = input.OpenSeekable())
        using (var memoryStream = new MemoryStream())
        {
            stream.CopyTo(memoryStream);
            docBytes = memoryStream.ToArray();
        }

        using (var ms = new MemoryStream(docBytes))
        using (var sDoc = SpreadsheetDocument.Open(ms, false))
        {
            if (sDoc.WorkbookPart is null)
            {
                throw CliException.InvalidFormat("Invalid spreadsheet. WorkbookPart is missing.");
            }

            if (!string.IsNullOrEmpty(tableName))
            {
                try
                {
                    rangeXml = SmlDataRetriever.RetrieveTable(sDoc, tableName);
                }
                catch (ArgumentException)
                {
                    throw CliException.InvalidArguments($"Excel table '{tableName}' was not found in the spreadsheet.");
                }
            }
            else
            {
                var sheets = SmlDataRetriever.SheetNames(sDoc);
                if (sheets.Length == 0)
                {
                    throw CliException.InvalidFormat("Invalid spreadsheet. No worksheets found.");
                }

                var targetSheet = sheetName;
                if (string.IsNullOrEmpty(targetSheet))
                {
                    targetSheet = sheets[0];
                }
                else if (!sheets.Contains(targetSheet))
                {
                    throw CliException.InvalidArguments($"Sheet '{targetSheet}' was not found in the spreadsheet.");
                }

                if (!string.IsNullOrEmpty(range))
                {
                    try
                    {
                        rangeXml = SmlDataRetriever.RetrieveRange(sDoc, targetSheet, range);
                    }
                    catch (Exception ex)
                    {
                        throw CliException.InvalidArguments(
                            $"Could not retrieve range '{range}' from sheet '{targetSheet}': {ex.Message}"
                        );
                    }
                }
                else
                {
                    try
                    {
                        rangeXml = SmlDataRetriever.RetrieveSheet(sDoc, targetSheet);
                    }
                    catch (Exception ex)
                    {
                        throw CliException.InvalidArguments($"Could not retrieve sheet '{targetSheet}': {ex.Message}");
                    }
                }
            }

            // Build converter settings
            var settings = new SmlToHtmlConverterSettings
            {
                PageTitle = pageTitle ?? input.LogicalName,
                AdditionalCss = additionalCss ?? string.Empty,
                CssClassPrefix = cssPrefix,
                FabricateCssClasses = fabricateCss,
            };

            // Convert using SmlToHtmlConverter.ConvertToHtml
            var htmlElement = SmlToHtmlConverter.ConvertToHtml(sDoc, settings, rangeXml);

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
    }
}
