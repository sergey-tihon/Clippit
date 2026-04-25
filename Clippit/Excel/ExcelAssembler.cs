// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Excel;

/// <summary>
/// Assembles an Excel document from a template by replacing <c>{{XPath}}</c> placeholders
/// with values resolved from a provided XML data element.
/// </summary>
/// <remarks>
/// <para>
/// Template cells may contain placeholder expressions of the form <c>{{xpath.expression}}</c>.
/// A single cell may contain multiple placeholders, and plain text surrounding placeholders
/// is preserved. For example, a cell containing <c>"Hello, {{Name}}!"</c> would produce
/// <c>"Hello, Alice!"</c> when the data element contains <c>&lt;Name&gt;Alice&lt;/Name&gt;</c>.
/// </para>
/// <para>
/// Placeholders are evaluated against the supplied <see cref="XElement"/> data root using
/// standard XPath 1.0 semantics. The data element is the XPath context node, so use
/// <em>relative</em> XPath expressions (e.g. <c>Name</c>, <c>./Orders/Order</c>,
/// <c>Item/@id</c>) rather than absolute paths starting with <c>/</c>. This is the same
/// convention used by <see cref="DocumentAssembler"/>. For node-set results, the text value
/// of the first matched node is used. For scalar results (string, number, boolean), the value
/// is converted to a string using <see cref="CultureInfo.InvariantCulture"/>. If the XPath
/// expression is invalid or throws an exception, the placeholder is replaced with
/// <c>[XPathError:expression]</c>.
/// </para>
/// <para>
/// Cells with resolved values are written back as inline strings, preserving the original
/// cell reference and any applied style index.
/// </para>
/// </remarks>
public static class ExcelAssembler
{
    private static readonly Regex s_placeholder = new(@"\{\{(.+?)\}\}", RegexOptions.Compiled);

    /// <summary>
    /// Assembles an Excel document from a template by substituting <c>{{XPath}}</c> placeholders
    /// with values resolved from the provided data element.
    /// </summary>
    /// <param name="template">The Excel template document containing <c>{{XPath}}</c> placeholders.</param>
    /// <param name="data">The XML data element used as the XPath evaluation context.</param>
    /// <returns>
    /// A new <see cref="SmlDocument"/> with the same file name as the template, with all
    /// placeholders replaced by their resolved values.
    /// </returns>
    public static SmlDocument AssembleDocument(SmlDocument template, XElement data)
    {
        var resultBytes = AssembleDocument(template.DocumentByteArray, data);
        return new SmlDocument(template.FileName, resultBytes);
    }

    /// <summary>
    /// Assembles an Excel document from a raw template byte array by substituting
    /// <c>{{XPath}}</c> placeholders with values resolved from the provided data element.
    /// </summary>
    /// <param name="templateBytes">The raw bytes of the Excel template (.xlsx).</param>
    /// <param name="data">The XML data element used as the XPath evaluation context.</param>
    /// <returns>A byte array containing the assembled Excel document.</returns>
    public static byte[] AssembleDocument(byte[] templateBytes, XElement data)
    {
        using var ms = new MemoryStream();
        ms.Write(templateBytes);
        ms.Position = 0;
        using (var doc = SpreadsheetDocument.Open(ms, isEditable: true))
        {
            AssembleInDocument(doc, data);
        }
        return ms.ToArray();
    }

    private static void AssembleInDocument(SpreadsheetDocument doc, XElement data)
    {
        var workbookPart = doc.WorkbookPart!;
        var sharedStrings = ReadSharedStrings(workbookPart);

        foreach (var worksheetPart in workbookPart.WorksheetParts)
        {
            var wsXDoc = worksheetPart.GetXDocument();
            var modified = false;

            foreach (var cell in wsXDoc.Descendants(S.c).ToList())
            {
                var cellText = GetCellText(cell, sharedStrings);
                if (cellText is null || !cellText.Contains("{{"))
                    continue;

                var resolved = s_placeholder.Replace(
                    cellText,
                    m =>
                    {
                        var xpath = m.Groups[1].Value.Trim();
                        try
                        {
                            var result = data.XPathEvaluate(xpath);
                            return result switch
                            {
                                IEnumerable seq when seq is not string => ResolveNodeSet(seq),
                                null => string.Empty,
                                _ => Convert.ToString(result, CultureInfo.InvariantCulture) ?? string.Empty,
                            };
                        }
                        catch (Exception ex)
                            when (ex is XPathException or ArgumentException or InvalidOperationException)
                        {
                            return $"[XPathError:{xpath}]";
                        }
                    }
                );

                // Rewrite cell value as an inline string, preserving non-value attributes/children.
                cell.SetAttributeValue(NoNamespace.t, "inlineStr");
                cell.Elements(S.v).Remove();
                cell.Elements(S._is).Remove();

                var inlineString = new XElement(S._is, new XElement(S.t, resolved));
                if (cell.Element(S.extLst) is { } extLst)
                    extLst.AddBeforeSelf(inlineString);
                else
                    cell.Add(inlineString);
                modified = true;
            }

            if (modified)
                worksheetPart.PutXDocument();
        }
    }

    private static string ResolveNodeSet(IEnumerable seq)
    {
        var first = seq.Cast<object?>().FirstOrDefault();
        return first switch
        {
            XElement xe => xe.Value,
            XAttribute xa => xa.Value,
            XText xt => xt.Value,
            null => string.Empty,
            _ => first.ToString() ?? string.Empty,
        };
    }

    private static string? GetCellText(XElement cell, IReadOnlyDictionary<int, string> sharedStrings)
    {
        var t = cell.Attribute(NoNamespace.t)?.Value;
        return t switch
        {
            "s" when int.TryParse(cell.Element(S.v)?.Value, out var idx) => sharedStrings.TryGetValue(idx, out var s)
                ? s
                : null,
            "inlineStr" => cell.Element(S._is)?.Element(S.t)?.Value,
            // "str" is used by SpreadsheetWriter for formula result string cells.
            "str" => cell.Element(S.v)?.Value,
            _ => null,
        };
    }

    private static IReadOnlyDictionary<int, string> ReadSharedStrings(WorkbookPart workbookPart)
    {
        var result = new Dictionary<int, string>();
        var ssPart = workbookPart.SharedStringTablePart;
        if (ssPart is null)
            return result;

        var items = ssPart.GetXDocument().Root?.Elements(S.si).ToList();
        if (items is null)
            return result;

        for (var i = 0; i < items.Count; i++)
        {
            // Shared string items may be a plain <t> child or multiple <r><t> rich-text runs.
            var text =
                items[i].Element(S.t)?.Value
                ?? string.Concat(items[i].Elements(S.r).Select(r => r.Element(S.t)?.Value ?? string.Empty));
            result[i] = text;
        }

        return result;
    }
}
