// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Excel
{
    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Global")]
    public class SmlToHtmlConverterSettings
    {
        public string PageTitle;
        public string CssClassPrefix;
        public bool FabricateCssClasses;
        public string GeneralCss;
        public string AdditionalCss;

        public SmlToHtmlConverterSettings()
        {
            PageTitle = "";
            CssClassPrefix = "pt-";
            FabricateCssClasses = true;
            GeneralCss = "span { white-space: pre-wrap; }";
            AdditionalCss = "";
        }

        public SmlToHtmlConverterSettings(SmlToHtmlConverterSettings htmlConverterSettings)
        {
            PageTitle = htmlConverterSettings.PageTitle;
            CssClassPrefix = htmlConverterSettings.CssClassPrefix;
            FabricateCssClasses = htmlConverterSettings.FabricateCssClasses;
            GeneralCss = htmlConverterSettings.GeneralCss;
            AdditionalCss = htmlConverterSettings.AdditionalCss;
        }
    }

    public static class SmlToHtmlConverter
    {
        // ***********************************************************************************************************************************
        #region PublicApis
        public static XElement ConvertTableToHtml(
            SmlDocument smlDoc,
            SmlToHtmlConverterSettings settings,
            string tableName
        )
        {
            using var ms = new MemoryStream();
            ms.Write(smlDoc.DocumentByteArray, 0, smlDoc.DocumentByteArray.Length);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            var rangeXml = SmlDataRetriever.RetrieveTable(sDoc, tableName);
            return ConvertToHtml(sDoc, settings, rangeXml);
        }

        public static XElement ConvertTableToHtml(
            SpreadsheetDocument sDoc,
            SmlToHtmlConverterSettings settings,
            string tableName
        )
        {
            var rangeXml = SmlDataRetriever.RetrieveTable(sDoc, tableName);
            return ConvertToHtml(sDoc, settings, rangeXml);
        }

        /// <summary>
        /// Converts a range/sheet/table XML element (produced by <see cref="SmlDataRetriever"/>) to an XHTML
        /// document element. The returned element is a fully-formed <c>&lt;html&gt;</c> tree with an embedded
        /// <c>&lt;style&gt;</c> block containing generated CSS classes (or inline <c>style</c> attributes when
        /// <see cref="SmlToHtmlConverterSettings.FabricateCssClasses"/> is <c>false</c>).
        /// </summary>
        public static XElement ConvertToHtml(
            SpreadsheetDocument sDoc,
            SmlToHtmlConverterSettings htmlConverterSettings,
            XElement rangeXml
        )
        {
            var xhtml = (XElement)ConvertToHtmlTransform(htmlConverterSettings, rangeXml);

            ReifyStylesAndClasses(htmlConverterSettings, xhtml);

            // Note: the xhtml returned by ConvertToHtmlTransform contains objects of type
            // XEntity.  PtOpenXmlUtil.cs define the XEntity class.  See
            // http://blogs.msdn.com/ericwhite/archive/2010/01/21/writing-entity-references-using-linq-to-xml.aspx
            // for detailed explanation.
            //
            // If you further transform the XML tree returned by ConvertToHtmlTransform, you
            // must do it correctly, or entities will not be serialized properly.

            return xhtml;
        }
        #endregion
        // ***********************************************************************************************************************************

        private static XNode ConvertToHtmlTransform(SmlToHtmlConverterSettings htmlConverterSettings, XNode node)
        {
            if (node is not XElement element)
                return node;

            // Root-level Data element from RetrieveSheet / RetrieveRange
            if (element.Name.LocalName == "Data")
                return BuildHtmlDocument(htmlConverterSettings, element);

            // Root-level Table element from RetrieveTable
            if (element.Name.LocalName == "Table")
                return BuildHtmlTable(htmlConverterSettings, element);

            // Recursively transform unknown elements (fallback)
            return new XElement(
                element.Name,
                element.Attributes(),
                element.Nodes().Select(n => ConvertToHtmlTransform(htmlConverterSettings, n))
            );
        }

        /// <summary>
        /// Builds an XHTML document from a &lt;Data&gt; element (sheet/range).
        /// </summary>
        private static XElement BuildHtmlDocument(
            SmlToHtmlConverterSettings htmlConverterSettings,
            XElement dataElement
        )
        {
            var dataProps = dataElement.Element("DataProps");

            // Extract column widths from <cols> inside DataProps.
            // Only present when the sheet has explicit column-width overrides; null otherwise.
            var columnWidths = ExtractColumnWidths(dataProps);

            var rows = dataElement.Elements("Row").Select(row => BuildTableRow(row, columnWidths));

            var thead = BuildTableHeader(columnWidths);

            var table = new XElement(
                Xhtml.table,
                new XAttribute("border", "1"),
                new XAttribute("cellpadding", "2"),
                new XAttribute("cellspacing", "0"),
                thead != null ? new XElement(Xhtml.thead, thead) : null,
                new XElement(Xhtml.tbody, rows)
            );

            var head = new XElement(
                Xhtml.head,
                new XElement(Xhtml.title, new XText(htmlConverterSettings.PageTitle)),
                new XElement(
                    Xhtml.meta,
                    new XAttribute("http-equiv", "Content-Type"),
                    new XAttribute("content", "text/html; charset=utf-8")
                )
            );

            return new XElement(Xhtml.html, head, new XElement(Xhtml.body, table));
        }

        /// <summary>
        /// Builds an XHTML document from a &lt;Table&gt; element (named table).
        /// </summary>
        private static XElement BuildHtmlTable(SmlToHtmlConverterSettings htmlConverterSettings, XElement tableElement)
        {
            var columns = tableElement.Element("Columns");
            var dataElement = tableElement.Element("Data");

            var tableName =
                (string)tableElement.Attribute("DisplayName") ?? (string)tableElement.Attribute("TableName");

            // Extract column names for header row
            var columnNames = columns?.Elements("Column").Select(c => (string)c.Attribute("Name")).ToList() ?? [];

            var columnWidths = ExtractColumnWidths(dataElement?.Element("DataProps"));

            // Build header row from column names
            var rows = new List<XElement>();
            if (columnNames.Count > 0)
            {
                var headerCells = columnNames.Select(name =>
                {
                    var td = new XElement(Xhtml.td, new XText(name ?? ""));
                    td.AddAnnotation(new Dictionary<string, string>());
                    return td;
                });
                rows.Add(new XElement(Xhtml.tr, headerCells));
            }

            if (dataElement != null)
            {
                foreach (var row in dataElement.Elements("Row"))
                    rows.Add(BuildTableRow(row, columnWidths));
            }

            var tableContent = new List<object>
            {
                new XAttribute("border", "1"),
                new XAttribute("cellpadding", "2"),
                new XAttribute("cellspacing", "0"),
            };

            if (tableName != null)
                tableContent.Add(new XElement(Xhtml.caption, new XText(tableName)));

            tableContent.Add(new XElement(Xhtml.tbody, rows));

            var head = new XElement(
                Xhtml.head,
                new XElement(Xhtml.title, new XText(htmlConverterSettings.PageTitle)),
                new XElement(
                    Xhtml.meta,
                    new XAttribute("http-equiv", "Content-Type"),
                    new XAttribute("content", "text/html; charset=utf-8")
                )
            );

            return new XElement(Xhtml.html, head, new XElement(Xhtml.body, new XElement(Xhtml.table, tableContent)));
        }

        /// <summary>
        /// Builds a &lt;tr&gt; from a &lt;Row&gt; element.
        /// </summary>
        private static XElement BuildTableRow(XElement rowElement, IReadOnlyList<double>? columnWidths)
        {
            var cells = rowElement.Elements("Cell").Select((cell, idx) => BuildTableCell(cell, idx, columnWidths));
            var tr = new XElement(Xhtml.tr, cells);

            var rowProps = rowElement.Element("RowProps");
            if (rowProps?.Attribute("ht") != null)
            {
                if (
                    double.TryParse(
                        (string)rowProps.Attribute("ht"),
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out var ht
                    )
                )
                {
                    tr.AddAnnotation(
                        new Dictionary<string, string>
                        {
                            ["height"] = ht.ToString("F2", CultureInfo.InvariantCulture) + "pt",
                        }
                    );
                }
            }

            return tr;
        }

        /// <summary>
        /// Builds a &lt;td&gt; from a &lt;Cell&gt; element, applying cell properties as CSS annotations.
        /// </summary>
        private static XElement BuildTableCell(
            XElement cellElement,
            int columnIndex,
            IReadOnlyList<double>? columnWidths
        )
        {
            var displayValue = (string)cellElement.Element("DisplayValue") ?? "";
            var cellProps = cellElement.Element("CellProps");

            var td = new XElement(Xhtml.td, new XText(displayValue));

            var styles = new Dictionary<string, string>();
            if (cellProps != null)
                ApplyCellStyles(cellProps, styles);

            if (columnWidths != null && columnIndex < columnWidths.Count && columnWidths[columnIndex] > 0)
                styles["width"] = columnWidths[columnIndex].ToString("F2", CultureInfo.InvariantCulture) + "pt";

            if (styles.Count > 0)
                td.AddAnnotation(styles);

            return td;
        }

        /// <summary>
        /// Builds a header row with Excel column letters (A, B, C …) sized according to
        /// <paramref name="columnWidths"/>. Returns <c>null</c> when no explicit column widths are
        /// available (most sheets without overrides won't have a header row).
        /// </summary>
        private static XElement? BuildTableHeader(IReadOnlyList<double>? columnWidths)
        {
            if (columnWidths == null || columnWidths.Count == 0)
                return null;

            var headerCells = columnWidths.Select(
                (_, idx) =>
                {
                    var th = new XElement(Xhtml.th, new XText(XlsxTables.IndexToColumnAddress(idx)));
                    th.AddAnnotation(new Dictionary<string, string>());
                    return th;
                }
            );

            return new XElement(Xhtml.tr, headerCells);
        }

        /// <summary>
        /// Extracts per-column widths (in points) from the &lt;cols&gt; element inside DataProps.
        /// Returns <c>null</c> when no explicit column-width overrides are present.
        /// </summary>
        private static List<double>? ExtractColumnWidths(XElement? dataProps)
        {
            var cols = dataProps?.Element("cols");
            if (cols == null)
                return null;

            var widths = new List<double>();
            foreach (var col in cols.Elements("col"))
            {
                var widthAttr = col.Attribute("width");
                widths.Add(
                    widthAttr != null
                    && double.TryParse((string)widthAttr, NumberStyles.Any, CultureInfo.InvariantCulture, out var w)
                        ? w * 7.5 // character widths → approximate points
                        : 0
                );
            }

            return widths.Count > 0 ? widths : null;
        }

        /// <summary>
        /// Applies cell-level CSS styles from CellProps into the styles dictionary.
        /// </summary>
        private static void ApplyCellStyles(XElement cellProps, Dictionary<string, string> styles)
        {
            var font = cellProps.Element("font");
            if (font != null)
            {
                if (font.Element("b") != null)
                    styles["font-weight"] = "bold";

                if (font.Element("i") != null)
                    styles["font-style"] = "italic";

                if (font.Element("u") != null)
                    styles["text-decoration"] = "underline";

                if (font.Element("strike") != null)
                    styles["text-decoration"] = "line-through";

                var sz = font.Element("sz");
                if (
                    sz?.Attribute("val") != null
                    && double.TryParse(
                        (string)sz.Attribute("val"),
                        NumberStyles.Any,
                        CultureInfo.InvariantCulture,
                        out var fontSize
                    )
                )
                    styles["font-size"] = fontSize.ToString("F1", CultureInfo.InvariantCulture) + "pt";

                var color = font.Element("color");
                if (color != null)
                {
                    var fontColor = ParseColor(color);
                    if (fontColor != null)
                        styles["color"] = fontColor;
                }

                var name = font.Element("name");
                if (name?.Attribute("val") != null)
                    CreateFontCssProperty((string)name.Attribute("val"), styles);

                var family = font.Element("family");
                if (family?.Attribute("val") != null)
                {
                    var genericFamily = ((string)family.Attribute("val")).ToLowerInvariant() switch
                    {
                        "roman" => "serif",
                        "swiss" => "sans-serif",
                        "modern" => "monospace",
                        "script" => "cursive",
                        "decorative" => "fantasy",
                        _ => null,
                    };
                    if (genericFamily != null && !styles.ContainsKey("font-family"))
                        styles["font-family"] = genericFamily;
                }
            }

            var fill = cellProps.Element("fill");
            if (fill != null)
            {
                var fgColor = fill.Element("patternFill")?.Element("fgColor");
                if (fgColor != null)
                {
                    var fillColor = ParseColor(fgColor);
                    if (fillColor != null)
                        styles["background-color"] = fillColor;
                }
            }

            var border = cellProps.Element("border");
            if (border != null)
                ApplyBorderStyles(border, styles);

            var alignment = cellProps.Element("alignment");
            if (alignment != null)
            {
                var horizontal = (string)alignment.Attribute("horizontal");
                if (horizontal != null)
                    styles["text-align"] = horizontal switch
                    {
                        "center" => "center",
                        "right" => "right",
                        "left" => "left",
                        "justify" => "justify",
                        _ => horizontal,
                    };

                var vertical = (string)alignment.Attribute("vertical");
                if (vertical != null)
                    styles["vertical-align"] = vertical switch
                    {
                        "top" => "top",
                        "center" => "middle",
                        "bottom" => "bottom",
                        _ => vertical,
                    };

                var wrapText = (string)alignment.Attribute("wrapText");
                if (wrapText == "1" || string.Equals(wrapText, "true", StringComparison.OrdinalIgnoreCase))
                    styles["white-space"] = "normal";

                var textRotation = (string)alignment.Attribute("textRotation");
                if (textRotation != null && int.TryParse(textRotation, out var rotation))
                {
                    // Only map the two Excel vertical-text cases to CSS writing-mode.
                    // rotation == 90: text rotated 90° counter-clockwise (vertical column)
                    // rotation == 255: stacked vertical text
                    // All other rotation values (e.g. 45°) are not representable in CSS
                    // table cells and are left unstyled to avoid unexpected layout.
                    if (rotation == 90 || rotation == 255)
                        styles["writing-mode"] = "vertical-rl";
                }

                var indent = (string)alignment.Attribute("indent");
                if (indent != null && int.TryParse(indent, out var indentLevel))
                    styles["padding-left"] = $"{indentLevel * 7}pt";
            }
        }

        private static void ApplyBorderStyles(XElement border, Dictionary<string, string> styles)
        {
            ApplySingleBorder(styles, "border-top", border.Element("top"));
            ApplySingleBorder(styles, "border-bottom", border.Element("bottom"));
            ApplySingleBorder(styles, "border-left", border.Element("left"));
            ApplySingleBorder(styles, "border-right", border.Element("right"));
        }

        private static void ApplySingleBorder(
            Dictionary<string, string> styles,
            string cssProperty,
            XElement? borderElement
        )
        {
            if (borderElement == null)
                return;

            var styleAttr = (string)borderElement.Attribute("style");
            if (styleAttr == null || styleAttr == "none")
                return;

            var cssBorderStyle = styleAttr switch
            {
                "double" => "double",
                "dotted" => "dotted",
                "dashed" or "mediumDashDot" or "dashDot" or "mediumDashDotDot" or "dashDotDot" or "slantDashDot" =>
                    "dashed",
                _ => "solid",
            };

            var cssBorderWidth = styleAttr switch
            {
                "medium" => "2px",
                "thick" => "3px",
                "hair" => "0.5px",
                "double" => "3px",
                _ => "1px",
            };

            var color = borderElement.Element("color");
            var cssColor = color != null ? ParseColor(color) ?? "#000000" : "#000000";

            styles[cssProperty] = $"{cssBorderWidth} {cssBorderStyle} {cssColor}";
        }

        /// <summary>
        /// Parses a color element and returns a CSS-compatible color string (e.g. #RRGGBB).
        /// Handles rgb, indexed, auto, and theme+tint color specifications.
        /// </summary>
        private static string? ParseColor(XElement colorElement)
        {
            // Direct ARGB value (e.g. "FF9C0006" — leading two hex digits are alpha, ignored)
            var rgb = (string)colorElement.Attribute("rgb");
            if (rgb != null)
            {
                return rgb.Length >= 8 ? "#" + rgb[2..] : "#" + rgb;
            }

            // auto="1" means "use the default foreground/background colour"
            if ((string)colorElement.Attribute("auto") == "1")
                return null;

            // Indexed colour palette
            var indexed = (string)colorElement.Attribute("indexed");
            if (
                indexed != null
                && int.TryParse(indexed, out var idx)
                && idx >= 0
                && idx < SmlDataRetriever.IndexedColors.Length
            )
            {
                var val = SmlDataRetriever.IndexedColors[idx];
                // Indexed values are stored as "00RRGGBB" — strip the leading "00"
                return val.Length >= 8 ? "#" + val[2..] : "#" + val;
            }

            // Theme colour with optional tint.
            // TODO: read the actual theme XML from sDoc.WorkbookPart.ThemePart for custom-themed workbooks;
            //       for now the default Office theme palette is used as an approximation.
            var theme = (string)colorElement.Attribute("theme");
            if (theme != null)
            {
                double tintValue = 0;
                var tintAttr = (string)colorElement.Attribute("tint");
                if (tintAttr != null)
                    double.TryParse(tintAttr, NumberStyles.Any, CultureInfo.InvariantCulture, out tintValue);

                var baseColor = GetThemeColor(int.TryParse(theme, out var ti) ? ti : 0);
                if (baseColor == null)
                    return "#000000";

                return Math.Abs(tintValue) > 0.001 ? ApplyTint(baseColor, tintValue) : baseColor;
            }

            return null;
        }

        /// <summary>
        /// Maps an Office theme colour index to its approximate hex value using the default Office theme.
        /// </summary>
        private static string? GetThemeColor(int themeIndex) =>
            themeIndex switch
            {
                0 => null, // none
                1 => "#FFFFFF", // background1 / white
                2 => "#000000", // text1 / black
                3 => "#808080", // background2 / gray
                4 => "#404040", // text2 / dark gray
                5 => "#4472C4", // accent1 / blue
                6 => "#ED7D31", // accent2 / orange
                7 => "#A5A5A5", // accent3 / gray
                8 => "#FFC000", // accent4 / gold
                9 => "#5B9BD5", // accent5 / light blue
                10 => "#70AD47", // accent6 / green
                11 => "#0563C1", // hyperlink
                12 => "#954F72", // followed hyperlink
                _ => null,
            };

        /// <summary>
        /// Lightens (positive tint) or darkens (negative tint) a hex colour.
        /// </summary>
        private static string ApplyTint(string hexColor, double tint)
        {
            try
            {
                var r = int.Parse(hexColor[1..3], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var g = int.Parse(hexColor[3..5], NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                var b = int.Parse(hexColor[5..7], NumberStyles.HexNumber, CultureInfo.InvariantCulture);

                int newR,
                    newG,
                    newB;
                if (tint >= 0)
                {
                    newR = (int)(r + tint * (255 - r));
                    newG = (int)(g + tint * (255 - g));
                    newB = (int)(b + tint * (255 - b));
                }
                else
                {
                    newR = (int)(r * (1 + tint));
                    newG = (int)(g * (1 + tint));
                    newB = (int)(b * (1 + tint));
                }

                return $"#{Math.Clamp(newR, 0, 255):X2}{Math.Clamp(newG, 0, 255):X2}{Math.Clamp(newB, 0, 255):X2}";
            }
            catch
            {
                return hexColor;
            }
        }

        private static void ReifyStylesAndClasses(SmlToHtmlConverterSettings htmlConverterSettings, XElement xhtml)
        {
            if (htmlConverterSettings.FabricateCssClasses)
            {
                var usedCssClassNames = new HashSet<string>();
                var elementsThatNeedClasses = xhtml
                    .DescendantsAndSelf()
                    .Select(d => new { Element = d, Styles = d.Annotation<Dictionary<string, string>>() })
                    .Where(z => z.Styles != null);
                var augmented = elementsThatNeedClasses
                    .Select(p => new
                    {
                        p.Element,
                        p.Styles,
                        StylesString = p.Element.Name.LocalName
                            + "|"
                            + p.Styles.OrderBy(k => k.Key).Select(s => $"{s.Key}: {s.Value};").StringConcatenate(),
                    })
                    .GroupBy(p => p.StylesString)
                    .ToList();
                var classCounter = 1000000;
                var sb = new StringBuilder();
                sb.Append(Environment.NewLine);
                foreach (var grp in augmented)
                {
                    string classNameToUse;
                    var firstOne = grp.First();
                    var styles = firstOne.Styles;
                    if (styles.ContainsKey("PtStyleName"))
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix + styles["PtStyleName"];
                        if (usedCssClassNames.Contains(classNameToUse))
                        {
                            classNameToUse =
                                htmlConverterSettings.CssClassPrefix
                                + styles["PtStyleName"]
                                + "-"
                                + classCounter.ToString().Substring(1);
                            classCounter++;
                        }
                    }
                    else
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix + classCounter.ToString().Substring(1);
                        classCounter++;
                    }
                    usedCssClassNames.Add(classNameToUse);
                    sb.Append(firstOne.Element.Name.LocalName + "." + classNameToUse + " {" + Environment.NewLine);
                    foreach (var st in firstOne.Styles.Where(s => s.Key != "PtStyleName"))
                    {
                        var s = "    " + st.Key + ": " + st.Value + ";" + Environment.NewLine;
                        sb.Append(s);
                    }
                    sb.Append("}" + Environment.NewLine);
                    var classAtt = new XAttribute("class", classNameToUse);
                    foreach (var gc in grp)
                        gc.Element.Add(classAtt);
                }
                var styleValue = htmlConverterSettings.GeneralCss + sb + htmlConverterSettings.AdditionalCss;

                SetStyleElementValue(xhtml, styleValue);
            }
            else
            {
                SetStyleElementValue(xhtml, htmlConverterSettings.GeneralCss + htmlConverterSettings.AdditionalCss);

                foreach (var d in xhtml.DescendantsAndSelf())
                {
                    var style = d.Annotation<Dictionary<string, string>>();
                    if (style == null)
                        continue;
                    var styleValue = style
                        .Where(p => p.Key != "PtStyleName")
                        .OrderBy(p => p.Key)
                        .Select(e => $"{e.Key}: {e.Value};")
                        .StringConcatenate();
                    var st = new XAttribute("style", styleValue);
                    if (d.Attribute("style") != null)
                        d.Attribute("style").Value += styleValue;
                    else
                        d.Add(st);
                }
            }
        }

        private static void SetStyleElementValue(XElement xhtml, string styleValue)
        {
            var styleElement = xhtml.Descendants(Xhtml.style).FirstOrDefault();
            if (styleElement != null)
                styleElement.Value = styleValue;
            else
            {
                styleElement = new XElement(Xhtml.style, styleValue);
                xhtml.Element(Xhtml.head)?.Add(styleElement);
            }
        }

        private static readonly Dictionary<string, string> FontFallback = new()
        {
            { "Arial", @"'{0}', sans-serif" },
            { "Arial Narrow", @"'{0}', sans-serif" },
            { "Arial Rounded MT Bold", @"'{0}', sans-serif" },
            { "Arial Unicode MS", @"'{0}', sans-serif" },
            { "Baskerville Old Face", @"'{0}', serif" },
            { "Berlin Sans FB", @"'{0}', sans-serif" },
            { "Berlin Sans FB Demi", @"'{0}', sans-serif" },
            { "Calibri Light", @"'{0}', sans-serif" },
            { "Gill Sans MT", @"'{0}', sans-serif" },
            { "Gill Sans MT Condensed", @"'{0}', sans-serif" },
            { "Lucida Sans", @"'{0}', sans-serif" },
            { "Lucida Sans Unicode", @"'{0}', sans-serif" },
            { "Segoe UI", @"'{0}', sans-serif" },
            { "Segoe UI Light", @"'{0}', sans-serif" },
            { "Segoe UI Semibold", @"'{0}', sans-serif" },
            { "Tahoma", @"'{0}', sans-serif" },
            { "Trebuchet MS", @"'{0}', sans-serif" },
            { "Verdana", @"'{0}', sans-serif" },
            { "Book Antiqua", @"'{0}', serif" },
            { "Bookman Old Style", @"'{0}', serif" },
            { "Californian FB", @"'{0}', serif" },
            { "Cambria", @"'{0}', serif" },
            { "Constantia", @"'{0}', serif" },
            { "Garamond", @"'{0}', serif" },
            { "Lucida Bright", @"'{0}', serif" },
            { "Lucida Fax", @"'{0}', serif" },
            { "Palatino Linotype", @"'{0}', serif" },
            { "Times New Roman", @"'{0}', serif" },
            { "Wide Latin", @"'{0}', serif" },
            { "Courier New", @"'{0}', monospace" },
            { "Lucida Console", @"'{0}', monospace" },
        };

        private static void CreateFontCssProperty(string font, Dictionary<string, string> style)
        {
            if (FontFallback.ContainsKey(font))
            {
                style.AddIfMissing("font-family", string.Format(FontFallback[font], font));
                return;
            }
            style.AddIfMissing("font-family", font);
        }
    }
}
