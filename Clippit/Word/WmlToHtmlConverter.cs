// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using SkiaSharp;

// 200e lrm - LTR
// 200f rlm - RTL

// todo need to set the HTTP "Content-Language" header, for instance:
// Content-Language: en-US
// Content-Language: fr-FR

namespace Clippit.Word
{
    public partial class WmlDocument
    {
        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public XElement ConvertToHtml(WmlToHtmlConverterSettings htmlConverterSettings)
        {
            return WmlToHtmlConverter.ConvertToHtml(this, htmlConverterSettings);
        }

        [SuppressMessage("ReSharper", "UnusedMember.Global")]
        public XElement ConvertToHtml(HtmlConverterSettings htmlConverterSettings)
        {
            var settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(this, settings);
        }
    }

    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Global")]
    public class WmlToHtmlConverterSettings
    {
        public string PageTitle;
        public string CssClassPrefix;
        public bool FabricateCssClasses;
        public string GeneralCss;
        public string AdditionalCss;
        public bool RestrictToSupportedLanguages;
        public bool RestrictToSupportedNumberingFormats;
        public Dictionary<string, Func<string, int, string, string>> ListItemImplementations;
        public Func<ImageInfo, XElement> ImageHandler;

        public WmlToHtmlConverterSettings()
        {
            PageTitle = "";
            CssClassPrefix = "pt-";
            FabricateCssClasses = true;
            GeneralCss = "span { white-space: pre-wrap; }";
            AdditionalCss = "";
            RestrictToSupportedLanguages = false;
            RestrictToSupportedNumberingFormats = false;
            ListItemImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations;
        }

        public WmlToHtmlConverterSettings(HtmlConverterSettings htmlConverterSettings)
        {
            PageTitle = htmlConverterSettings.PageTitle;
            CssClassPrefix = htmlConverterSettings.CssClassPrefix;
            FabricateCssClasses = htmlConverterSettings.FabricateCssClasses;
            GeneralCss = htmlConverterSettings.GeneralCss;
            AdditionalCss = htmlConverterSettings.AdditionalCss;
            RestrictToSupportedLanguages = htmlConverterSettings.RestrictToSupportedLanguages;
            RestrictToSupportedNumberingFormats = htmlConverterSettings.RestrictToSupportedNumberingFormats;
            ListItemImplementations = htmlConverterSettings.ListItemImplementations;
            ImageHandler = htmlConverterSettings.ImageHandler;
        }
    }

    [SuppressMessage("ReSharper", "FieldCanBeMadeReadOnly.Global")]
    public class HtmlConverterSettings
    {
        public string PageTitle = "";
        public string CssClassPrefix = "pt-";
        public bool FabricateCssClasses = true;
        public string GeneralCss = "span { white-space: pre-wrap; }";
        public string AdditionalCss = "";
        public bool RestrictToSupportedLanguages = false;
        public bool RestrictToSupportedNumberingFormats = false;
        public Dictionary<string, Func<string, int, string, string>> ListItemImplementations =
            ListItemRetrieverSettings.DefaultListItemTextImplementations;
        public Func<ImageInfo, XElement> ImageHandler;
    }

    public static class HtmlConverter
    {
        public static XElement ConvertToHtml(WmlDocument wmlDoc, HtmlConverterSettings htmlConverterSettings)
        {
            var settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(wmlDoc, settings);
        }

        public static XElement ConvertToHtml(WordprocessingDocument wDoc, HtmlConverterSettings htmlConverterSettings)
        {
            var settings = new WmlToHtmlConverterSettings(htmlConverterSettings);
            return WmlToHtmlConverter.ConvertToHtml(wDoc, settings);
        }
    }

    [SuppressMessage("ReSharper", "NotAccessedField.Global")]
    [SuppressMessage("ReSharper", "UnusedMember.Global")]
    public class ImageInfo
    {
        public SKBitmap Image;
        public XAttribute ImgStyleAttribute;
        public string ContentType;
        public XElement DrawingElement;
        public string AltText;

        public const int EmusPerInch = 914400;
        public const int EmusPerCm = 360000;
    }

    public static class WmlToHtmlConverter
    {
        public static XElement ConvertToHtml(WmlDocument doc, WmlToHtmlConverterSettings htmlConverterSettings)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(doc);
            using var document = streamDoc.GetWordprocessingDocument();
            return ConvertToHtml(document, htmlConverterSettings);
        }

        public static XElement ConvertToHtml(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings htmlConverterSettings
        )
        {
            RevisionAccepter.AcceptRevisions(wordDoc);
            var simplifyMarkupSettings = new SimplifyMarkupSettings
            {
                RemoveComments = true,
                RemoveContentControls = true,
                RemoveEndAndFootNotes = true,
                RemoveFieldCodes = false,
                RemoveLastRenderedPageBreak = true,
                RemovePermissions = true,
                RemoveProof = true,
                RemoveRsidInfo = true,
                RemoveSmartTags = true,
                RemoveSoftHyphens = true,
                RemoveGoBackBookmark = true,
                ReplaceTabsWithSpaces = false,
            };
            MarkupSimplifier.SimplifyMarkup(wordDoc, simplifyMarkupSettings);

            var formattingAssemblerSettings = new FormattingAssemblerSettings
            {
                RemoveStyleNamesFromParagraphAndRunProperties = false,
                ClearStyles = false,
                RestrictToSupportedLanguages = htmlConverterSettings.RestrictToSupportedLanguages,
                RestrictToSupportedNumberingFormats = htmlConverterSettings.RestrictToSupportedNumberingFormats,
                CreateHtmlConverterAnnotationAttributes = true,
                OrderElementsPerStandard = false,
                ListItemRetrieverSettings = htmlConverterSettings.ListItemImplementations is null
                    ? new ListItemRetrieverSettings()
                    {
                        ListItemTextImplementations = ListItemRetrieverSettings.DefaultListItemTextImplementations,
                    }
                    : new ListItemRetrieverSettings()
                    {
                        ListItemTextImplementations = htmlConverterSettings.ListItemImplementations,
                    },
            };

            FormattingAssembler.AssembleFormatting(wordDoc, formattingAssemblerSettings);

            InsertAppropriateNonbreakingSpaces(wordDoc);
            CalculateSpanWidthForTabs(wordDoc);
            ReverseTableBordersForRtlTables(wordDoc);
            AdjustTableBorders(wordDoc);
            var rootElement = wordDoc.MainDocumentPart.GetXDocument().Root;
            FieldRetriever.AnnotateWithFieldInfo(wordDoc.MainDocumentPart);
            AnnotateForSections(wordDoc);

            var xhtml = (XElement)ConvertToHtmlTransform(wordDoc, htmlConverterSettings, rootElement, false, 0m);

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

        private static void ReverseTableBordersForRtlTables(WordprocessingDocument wordDoc)
        {
            var xd = wordDoc.MainDocumentPart.GetXDocument();
            foreach (var tbl in xd.Descendants(W.tbl))
            {
                var bidiVisual = tbl.Elements(W.tblPr).Elements(W.bidiVisual).FirstOrDefault();
                if (bidiVisual is null)
                    continue;

                var tblBorders = tbl.Elements(W.tblPr).Elements(W.tblBorders).FirstOrDefault();
                if (tblBorders is not null)
                {
                    var left = tblBorders.Element(W.left);
                    if (left is not null)
                        left = new XElement(W.right, left.Attributes());

                    var right = tblBorders.Element(W.right);
                    if (right is not null)
                        right = new XElement(W.left, right.Attributes());

                    var newTblBorders = new XElement(
                        W.tblBorders,
                        tblBorders.Element(W.top),
                        left,
                        tblBorders.Element(W.bottom),
                        right
                    );
                    tblBorders.ReplaceWith(newTblBorders);
                }

                foreach (var tc in tbl.Elements(W.tr).Elements(W.tc))
                {
                    var tcBorders = tc.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault();
                    if (tcBorders is not null)
                    {
                        var left = tcBorders.Element(W.left);
                        if (left is not null)
                            left = new XElement(W.right, left.Attributes());

                        var right = tcBorders.Element(W.right);
                        if (right is not null)
                            right = new XElement(W.left, right.Attributes());

                        var newTcBorders = new XElement(
                            W.tcBorders,
                            tcBorders.Element(W.top),
                            left,
                            tcBorders.Element(W.bottom),
                            right
                        );
                        tcBorders.ReplaceWith(newTcBorders);
                    }
                }
            }
        }

        private static void ReifyStylesAndClasses(WmlToHtmlConverterSettings htmlConverterSettings, XElement xhtml)
        {
            if (htmlConverterSettings.FabricateCssClasses)
            {
                var usedCssClassNames = new HashSet<string>();
                var elementsThatNeedClasses = xhtml
                    .DescendantsAndSelf()
                    .Select(d => new { Element = d, Styles = d.Annotation<Dictionary<string, string>>() })
                    .Where(z => z.Styles is not null);
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
                    if (styles.TryGetValue("PtStyleName", out var ptStyleName))
                    {
                        classNameToUse = htmlConverterSettings.CssClassPrefix + ptStyleName;
                        if (usedCssClassNames.Contains(classNameToUse))
                        {
                            classNameToUse =
                                htmlConverterSettings.CssClassPrefix
                                + ptStyleName
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
                // Previously, the h:style element was not added at this point. However,
                // at least the General CSS will contain important settings.
                SetStyleElementValue(xhtml, htmlConverterSettings.GeneralCss + htmlConverterSettings.AdditionalCss);

                foreach (var d in xhtml.DescendantsAndSelf())
                {
                    var style = d.Annotation<Dictionary<string, string>>();
                    if (style is null)
                        continue;
                    var styleValue = style
                        .Where(p => p.Key != "PtStyleName")
                        .OrderBy(p => p.Key)
                        .Select(e => $"{e.Key}: {e.Value};")
                        .StringConcatenate();
                    var st = new XAttribute("style", styleValue);
                    if (d.Attribute("style") is not null)
                        d.Attribute("style").Value += styleValue;
                    else
                        d.Add(st);
                }
            }
        }

        private static void SetStyleElementValue(XElement xhtml, string styleValue)
        {
            var styleElement = xhtml.Descendants(Xhtml.style).FirstOrDefault();
            if (styleElement is not null)
                styleElement.Value = styleValue;
            else
            {
                styleElement = new XElement(Xhtml.style, styleValue);
                var head = xhtml.Element(Xhtml.head);
                if (head is not null)
                    head.Add(styleElement);
            }
        }

        private static object ConvertToHtmlTransform(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XNode node,
            bool suppressTrailingWhiteSpace,
            decimal currentMarginLeft
        )
        {
            if (node is not XElement element)
                return null;

            // Transform the w:document element to the XHTML h:html element.
            // The h:head element is laid out based on the W3C's recommended layout, i.e.,
            // the charset (using the HTML5-compliant form), the title (which is always
            // there but possibly empty), and other meta tags.
            if (element.Name == W.document)
            {
                return new XElement(
                    Xhtml.html,
                    new XElement(
                        Xhtml.head,
                        new XElement(Xhtml.meta, new XAttribute("charset", "UTF-8")),
                        settings.PageTitle is not null
                            ? new XElement(Xhtml.title, new XText(settings.PageTitle))
                            : new XElement(Xhtml.title, new XText(string.Empty)),
                        new XElement(
                            Xhtml.meta,
                            new XAttribute("name", "Generator"),
                            new XAttribute("content", "PowerTools for Open XML")
                        )
                    ),
                    element
                        .Elements()
                        .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
                );
            }

            // Transform the w:body element to the XHTML h:body element.
            if (element.Name == W.body)
            {
                return new XElement(Xhtml.body, CreateSectionDivs(wordDoc, settings, element));
            }

            // Transform the w:p element to the XHTML h:h1-h6 or h:p element (if the previous paragraph does not
            // have a style separator).
            if (element.Name == W.p)
            {
                return ProcessParagraph(wordDoc, settings, element, suppressTrailingWhiteSpace, currentMarginLeft);
            }

            // Transform hyperlinks to the XHTML h:a element.
            if (element.Name == W.hyperlink && element.Attribute(R.id) is not null)
            {
                try
                {
                    var a = new XElement(
                        Xhtml.a,
                        new XAttribute(
                            "href",
                            wordDoc
                                .MainDocumentPart.HyperlinkRelationships.First(x =>
                                    x.Id == (string)element.Attribute(R.id)
                                )
                                .Uri
                        ),
                        element.Elements(W.r).Select(run => ConvertRun(wordDoc, settings, run))
                    );
                    if (!a.Nodes().Any())
                        a.Add(new XText(""));
                    return a;
                }
                catch (UriFormatException)
                {
                    return element
                        .Elements()
                        .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
                }
            }

            // Transform hyperlinks to bookmarks to the XHTML h:a element.
            if (element.Name == W.hyperlink && element.Attribute(W.anchor) is not null)
            {
                return ProcessHyperlinkToBookmark(wordDoc, settings, element);
            }

            // Transform contents of runs.
            if (element.Name == W.r)
            {
                return ConvertRun(wordDoc, settings, element);
            }

            // Transform w:bookmarkStart into anchor
            if (element.Name == W.bookmarkStart)
            {
                return ProcessBookmarkStart(element);
            }

            // Transform every w:t element to a text node.
            if (element.Name == W.t)
            {
                // We don't need to convert characters to entities in a UTF-8 document.
                // Further, we don't need &nbsp; entities for significant whitespace
                // because we are wrapping the text nodes in <span> elements within
                // which all whitespace is significant.
                return new XText(element.Value);
            }

            // Transform symbols to spans
            if (element.Name == W.sym)
            {
                var cs = (string)element.Attribute(W._char);
                var c = Convert.ToInt32(cs, 16);
                return new XElement(Xhtml.span, new XEntity($"#{c}"));
            }

            // Transform tabs that have the pt:TabWidth attribute set
            if (element.Name == W.tab)
            {
                return ProcessTab(element);
            }

            // Transform w:br to h:br.
            if (element.Name == W.br || element.Name == W.cr)
            {
                return ProcessBreak(element);
            }

            // Transform w:noBreakHyphen to '-'
            if (element.Name == W.noBreakHyphen)
            {
                return new XText("-");
            }

            // Transform w:tbl to h:tbl.
            if (element.Name == W.tbl)
            {
                return ProcessTable(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform w:tr to h:tr.
            if (element.Name == W.tr)
            {
                return ProcessTableRow(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform w:tc to h:td.
            if (element.Name == W.tc)
            {
                return ProcessTableCell(wordDoc, settings, element);
            }

            // Transform images and text boxes.
            if (element.Name == W.drawing || element.Name == W.pict || element.Name == W._object)
            {
                // Text boxes in w:drawing (wps:wsp/wps:txbx) must be handled before image processing.
                if (element.Name == W.drawing)
                {
                    var textBoxResult = ProcessTextBoxDrawing(wordDoc, settings, element);
                    if (textBoxResult is not null)
                        return textBoxResult;
                }
                return ProcessImage(wordDoc, element, settings.ImageHandler);
            }

            // Transform content controls.
            if (element.Name == W.sdt)
            {
                return ProcessContentControl(wordDoc, settings, element, currentMarginLeft);
            }

            // Transform smart tags and simple fields.
            if (element.Name == W.smartTag || element.Name == W.fldSimple)
            {
                return CreateBorderDivs(wordDoc, settings, element.Elements());
            }

            // Ignore element.
            return null;
        }

        private static object ProcessHyperlinkToBookmark(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement element
        )
        {
            var style = new Dictionary<string, string>();
            var a = new XElement(
                Xhtml.a,
                new XAttribute("href", "#" + (string)element.Attribute(W.anchor)),
                element.Elements(W.r).Select(run => ConvertRun(wordDoc, settings, run))
            );
            if (!a.Nodes().Any())
                a.Add(new XText(""));
            style.Add("text-decoration", "none");
            a.AddAnnotation(style);
            return a;
        }

        private static object ProcessBookmarkStart(XElement element)
        {
            var name = (string)element.Attribute(W.name);
            if (name is null)
                return null;

            var style = new Dictionary<string, string>();
            var a = new XElement(Xhtml.a, new XAttribute("id", name), new XText(""));
            if (!a.Nodes().Any())
                a.Add(new XText(""));
            style.Add("text-decoration", "none");
            a.AddAnnotation(style);
            return a;
        }

        private static object ProcessTab(XElement element)
        {
            var tabWidthAtt = element.Attribute(PtOpenXml.TabWidth);
            if (tabWidthAtt is null)
                return null;

            var leader = (string)element.Attribute(PtOpenXml.Leader);
            var tabWidth = (decimal)tabWidthAtt;
            var style = new Dictionary<string, string>();
            XElement span;
            if (leader is not null)
            {
                var leaderChar = leader switch
                {
                    "hyphen" => "-",
                    "dot" => ".",
                    "underscore" => "_",
                    _ => ".",
                };

                var runContainingTabToReplace = element.Ancestors(W.r).First();
                var fontNameAtt =
                    runContainingTabToReplace.Attribute(PtOpenXml.pt + "FontName")
                    ?? runContainingTabToReplace.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");

                var dummyRun = new XElement(
                    W.r,
                    fontNameAtt,
                    runContainingTabToReplace.Elements(W.rPr),
                    new XElement(W.t, leaderChar)
                );

                var widthOfLeaderChar = WordprocessingMLUtil.CalcWidthOfRunInTwips(dummyRun);

                var forceArial = false;
                if (widthOfLeaderChar == 0)
                {
                    dummyRun = new XElement(
                        W.r,
                        new XAttribute(PtOpenXml.FontName, "Arial"),
                        runContainingTabToReplace.Elements(W.rPr),
                        new XElement(W.t, leaderChar)
                    );
                    widthOfLeaderChar = WordprocessingMLUtil.CalcWidthOfRunInTwips(dummyRun);
                    forceArial = true;
                }

                if (widthOfLeaderChar != 0)
                {
                    var numberOfLeaderChars = (int)(Math.Floor((tabWidth * 1440) / widthOfLeaderChar));
                    if (numberOfLeaderChars < 0)
                        numberOfLeaderChars = 0;
                    span = new XElement(
                        Xhtml.span,
                        new XAttribute(XNamespace.Xml + "space", "preserve"),
                        " " + "".PadRight(numberOfLeaderChars, leaderChar[0]) + " "
                    );
                    style.Add("margin", "0 0 0 0");
                    style.Add("padding", "0 0 0 0");
                    style.Add("width", FormattableString.Invariant($"{tabWidth:0.00}in"));
                    style.Add("text-align", "center");
                    if (forceArial)
                        style.Add("font-family", "Arial");
                }
                else
                {
                    span = new XElement(Xhtml.span, new XAttribute(XNamespace.Xml + "space", "preserve"), " ");
                    style.Add("margin", "0 0 0 0");
                    style.Add("padding", "0 0 0 0");
                    style.Add("width", FormattableString.Invariant($"{tabWidth:0.00}in"));
                    style.Add("text-align", "center");
                    if (leader == "underscore")
                    {
                        style.Add("text-decoration", "underline");
                    }
                }
            }
            else
            {
#if false
                            var bidi = element
                                .Ancestors(W.p)
                                .Take(1)
                                .Elements(W.pPr)
                                .Elements(W.bidi)
                                .Where(b => b.Attribute(W.val) is null || b.Attribute(W.val).ToBoolean() == true)
                                .FirstOrDefault();
                            var isBidi = bidi is not null;
                            if (isBidi)
                                span = new XElement(Xhtml.span, new XEntity("#x200f")); // RLM
                            else
                                span = new XElement(Xhtml.span, new XEntity("#x200e")); // LRM
#else
                span = new XElement(Xhtml.span, new XEntity("#x00a0"));
#endif
                style.Add("margin", FormattableString.Invariant($"0 0 0 {tabWidth:0.00}in"));
                style.Add("padding", "0 0 0 0");
            }
            span.AddAnnotation(style);
            return span;
        }

        private static object ProcessBreak(XElement element)
        {
            var breakType = (string)element.Attribute(W.type);

            // Page and column breaks are rendered as a print-CSS page-break marker.
            if (breakType is "page" or "column")
            {
                var pageBreakSpan = new XElement(Xhtml.span);
                pageBreakSpan.AddAnnotation(
                    new Dictionary<string, string> { { "display", "block" }, { "page-break-before", "always" } }
                );
                return pageBreakSpan;
            }

            XElement span = null;
            var tabWidth = (decimal?)element.Attribute(PtOpenXml.TabWidth);
            if (tabWidth is not null)
            {
                span = new XElement(Xhtml.span);
                span.AddAnnotation(
                    new Dictionary<string, string>
                    {
                        { "margin", FormattableString.Invariant($"0 0 0 {tabWidth:0.00}in") },
                        { "padding", "0 0 0 0" },
                    }
                );
            }

            var paragraph = element.Ancestors(W.p).FirstOrDefault();
            var isBidi =
                paragraph is not null
                && paragraph
                    .Elements(W.pPr)
                    .Elements(W.bidi)
                    .Any(b => b.Attribute(W.val) is null || b.Attribute(W.val).ToBoolean() == true);
            var zeroWidthChar = isBidi ? new XEntity("#x200f") : new XEntity("#x200e");

            return new object[] { new XElement(Xhtml.br), zeroWidthChar, span };
        }

        private static object ProcessContentControl(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement element,
            decimal currentMarginLeft
        )
        {
            var relevantAncestors = element.Ancestors().TakeWhile(a => a.Name != W.txbxContent);
            var isRunLevelContentControl = relevantAncestors.Any(a => a.Name == W.p);
            if (isRunLevelContentControl)
            {
                return element
                    .Elements(W.sdtContent)
                    .Elements()
                    .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
                    .ToList();
            }
            return CreateBorderDivs(wordDoc, settings, element.Elements(W.sdtContent).Elements());
        }

        // Transform the w:p element, including the following sibling w:p element(s)
        // in case the w:p element has a style separator. The sibling(s) will be
        // transformed to h:span elements rather than h:p elements and added to
        // the element (e.g., h:h2) created from the w:p element having the (first)
        // style separator (i.e., a w:specVanish element).
        private static object ProcessParagraph(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement element,
            bool suppressTrailingWhiteSpace,
            decimal currentMarginLeft
        )
        {
            // Ignore this paragraph if the previous paragraph has a style separator.
            // We have already transformed this one together with the previous one.
            var previousParagraph = element.ElementsBeforeSelf(W.p).LastOrDefault();
            if (HasStyleSeparator(previousParagraph))
                return null;

            var elementName = GetParagraphElementName(element, wordDoc);
            var isBidi = IsBidi(element);
            var paragraph = (XElement)ConvertParagraph(
                wordDoc,
                settings,
                element,
                elementName,
                suppressTrailingWhiteSpace,
                currentMarginLeft,
                isBidi
            );

            // The paragraph conversion might have created empty spans.
            // These can and should be removed because empty spans are
            // invalid in HTML5.
            paragraph
                .Elements(Xhtml.span)
                .Where(e =>
                    e.IsEmpty
                    && (
                        e.Annotation<Dictionary<string, string>>() is not { } style
                        || !style.ContainsKey("page-break-before")
                    )
                )
                .Remove();

            foreach (var span in paragraph.Elements(Xhtml.span).ToList())
            {
                var v = span.Value;
                if (
                    v.Length > 0
                    && (char.IsWhiteSpace(v[0]) || char.IsWhiteSpace(v[v.Length - 1]))
                    && span.Attribute(XNamespace.Xml + "space") is null
                )
                    span.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
            }

            while (HasStyleSeparator(element))
            {
                element = element.ElementsAfterSelf(W.p).FirstOrDefault();
                if (element is null)
                    break;

                elementName = Xhtml.span;
                isBidi = IsBidi(element);
                var span = (XElement)ConvertParagraph(
                    wordDoc,
                    settings,
                    element,
                    elementName,
                    suppressTrailingWhiteSpace,
                    currentMarginLeft,
                    isBidi
                );
                var v = span.Value;
                if (
                    v.Length > 0
                    && (char.IsWhiteSpace(v[0]) || char.IsWhiteSpace(v[v.Length - 1]))
                    && span.Attribute(XNamespace.Xml + "space") is null
                )
                    span.Add(new XAttribute(XNamespace.Xml + "space", "preserve"));
                paragraph.Add(span);
            }

            return paragraph;
        }

        private static object ProcessTable(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement element,
            decimal currentMarginLeft
        )
        {
            var style = new Dictionary<string, string>();
            style.AddIfMissing("border-collapse", "collapse");
            style.AddIfMissing("border", "none");
            var bidiVisual = element.Elements(W.tblPr).Elements(W.bidiVisual).FirstOrDefault();
            var tblW = element.Elements(W.tblPr).Elements(W.tblW).FirstOrDefault();
            if (tblW is not null)
            {
                var type = (string)tblW.Attribute(W.type);
                if (type is not null && type == "pct")
                {
                    var w = (int)tblW.Attribute(W._w);
                    style.AddIfMissing("width", (w / 50) + "%");
                }
            }
            var tblInd = element.Elements(W.tblPr).Elements(W.tblInd).FirstOrDefault();
            if (tblInd is not null)
            {
                var tblIndType = (string)tblInd.Attribute(W.type);
                if (tblIndType is not null)
                {
                    if (tblIndType == "dxa")
                    {
                        var width = (decimal?)tblInd.Attribute(W._w);
                        if (width is not null)
                        {
                            style.AddIfMissing(
                                "margin-left",
                                width > 0m ? FormattableString.Invariant($"{width / 20m}pt") : "0"
                            );
                        }
                    }
                }
            }
            var tableDirection = bidiVisual is not null ? new XAttribute("dir", "rtl") : new XAttribute("dir", "ltr");
            style.AddIfMissing("margin-bottom", ".001pt");

            // Handle floating table (w:tblpPr): apply CSS float/margin so surrounding text flows around the table.
            var tblpPr = element.Elements(W.tblPr).Elements(W.tblpPr).FirstOrDefault();
            Dictionary<string, string>? wrapperDivStyle = null;
            if (tblpPr is not null)
            {
                // Map w:tblpXSpec to a CSS float value. CSS float has no clean equivalent for "center",
                // and absolute positioning (w:tblpX/w:tblpY) cannot be expressed with float at all —
                // in those cases we intentionally omit float but still honor the *FromText margins so
                // the table at least renders with appropriate spacing.
                var xSpec = (string)tblpPr.Attribute(W.tblpXSpec);
                var floatValue = xSpec switch
                {
                    "left" => "left",
                    "right" => "right",
                    _ => null,
                };
                if (floatValue is not null)
                {
                    wrapperDivStyle ??= new Dictionary<string, string>();
                    wrapperDivStyle["float"] = floatValue;
                }

                static string? TwipsToPoints(XAttribute attr) =>
                    attr is not null && decimal.TryParse((string)attr, out var v)
                        ? FormattableString.Invariant($"{v / 20m:0.##}pt")
                        : null;

                var marginLeft = TwipsToPoints(tblpPr.Attribute(W.leftFromText));
                var marginRight = TwipsToPoints(tblpPr.Attribute(W.rightFromText));
                var marginTop = TwipsToPoints(tblpPr.Attribute(W.topFromText));
                var marginBottom = TwipsToPoints(tblpPr.Attribute(W.bottomFromText));

                if (marginLeft is not null)
                {
                    wrapperDivStyle ??= new Dictionary<string, string>();
                    wrapperDivStyle["margin-left"] = marginLeft;
                }
                if (marginRight is not null)
                {
                    wrapperDivStyle ??= new Dictionary<string, string>();
                    wrapperDivStyle["margin-right"] = marginRight;
                }
                if (marginTop is not null)
                {
                    wrapperDivStyle ??= new Dictionary<string, string>();
                    wrapperDivStyle["margin-top"] = marginTop;
                }
                if (marginBottom is not null)
                {
                    wrapperDivStyle ??= new Dictionary<string, string>();
                    wrapperDivStyle["margin-bottom"] = marginBottom;
                }
            }

            var table = new XElement(
                Xhtml.table,
                // TODO: Revisit and make sure the omission is covered by appropriate CSS.
                // new XAttribute("border", "1"),
                // new XAttribute("cellspacing", 0),
                // new XAttribute("cellpadding", 0),
                tableDirection,
                element.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
            );
            table.AddAnnotation(style);
            var jc = (string)element.Elements(W.tblPr).Elements(W.jc).Attributes(W.val).FirstOrDefault() ?? "left";
            XAttribute dir = null;
            XAttribute jcToUse = null;
            if (bidiVisual is not null)
            {
                dir = new XAttribute("dir", "rtl");
                jcToUse = jc switch
                {
                    "left" => new XAttribute("align", "right"),
                    "right" => new XAttribute("align", "left"),
                    "center" => new XAttribute("align", "center"),
                    _ => jcToUse,
                };
            }
            else
            {
                jcToUse = new XAttribute("align", jc);
            }
            var tableDiv = new XElement(Xhtml.div, dir, jcToUse, table);
            if (wrapperDivStyle is not null)
            {
                tableDiv.AddAnnotation(wrapperDivStyle);
            }
            return tableDiv;
        }

        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        private static object ProcessTableCell(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement element
        )
        {
            var style = new Dictionary<string, string>();
            XAttribute colSpan = null;
            XAttribute rowSpan = null;

            var tcPr = element.Element(W.tcPr);
            if (tcPr is not null)
            {
                if ((string)tcPr.Elements(W.vMerge).Attributes(W.val).FirstOrDefault() == "restart")
                {
                    var currentRow = element.Parent.ElementsBeforeSelf(W.tr).Count();
                    var currentCell = element.ElementsBeforeSelf(W.tc).Count();
                    var tbl = element.Parent.Parent;
                    var rowSpanCount = 1;
                    currentRow += 1;
                    while (true)
                    {
                        var row = tbl.Elements(W.tr).Skip(currentRow).FirstOrDefault();
                        if (row is null)
                            break;
                        var cell2 = row.Elements(W.tc).Skip(currentCell).FirstOrDefault();
                        if (cell2 is null)
                            break;
                        if (cell2.Elements(W.tcPr).Elements(W.vMerge).FirstOrDefault() is null)
                            break;
                        if (
                            (string)cell2.Elements(W.tcPr).Elements(W.vMerge).Attributes(W.val).FirstOrDefault()
                            == "restart"
                        )
                            break;
                        currentRow += 1;
                        rowSpanCount += 1;
                    }
                    rowSpan = new XAttribute("rowspan", rowSpanCount);
                }

                if (
                    tcPr.Element(W.vMerge) is not null
                    && (string)tcPr.Elements(W.vMerge).Attributes(W.val).FirstOrDefault() != "restart"
                )
                    return null;

                if (tcPr.Element(W.vAlign) is not null)
                {
                    var vAlignVal = (string)tcPr.Elements(W.vAlign).Attributes(W.val).FirstOrDefault();
                    if (vAlignVal == "top")
                        style.AddIfMissing("vertical-align", "top");
                    else if (vAlignVal == "center")
                        style.AddIfMissing("vertical-align", "middle");
                    else if (vAlignVal == "bottom")
                        style.AddIfMissing("vertical-align", "bottom");
                    else
                        style.AddIfMissing("vertical-align", "middle");
                }
                style.AddIfMissing("vertical-align", "top");

                if ((string)tcPr.Elements(W.tcW).Attributes(W.type).FirstOrDefault() == "dxa")
                {
                    decimal width = (int)tcPr.Elements(W.tcW).Attributes(W._w).FirstOrDefault();
                    style.AddIfMissing("width", FormattableString.Invariant($"{width / 20m}pt"));
                }
                if ((string)tcPr.Elements(W.tcW).Attributes(W.type).FirstOrDefault() == "pct")
                {
                    decimal width = (int)tcPr.Elements(W.tcW).Attributes(W._w).FirstOrDefault();
                    style.AddIfMissing("width", FormattableString.Invariant($"{width / 50m:0.0}%"));
                }

                var tcBorders = tcPr.Element(W.tcBorders);
                GenerateBorderStyle(tcBorders, W.top, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.right, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.bottom, style, BorderType.Cell);
                GenerateBorderStyle(tcBorders, W.left, style, BorderType.Cell);

                CreateStyleFromShd(style, tcPr.Element(W.shd));

                var gridSpan = tcPr.Elements(W.gridSpan).Attributes(W.val).Select(a => (int?)a).FirstOrDefault();
                if (gridSpan is not null)
                    colSpan = new XAttribute("colspan", (int)gridSpan);
            }
            style.AddIfMissing("padding-top", "0");
            style.AddIfMissing("padding-bottom", "0");

            var cell = new XElement(
                Xhtml.td,
                rowSpan,
                colSpan,
                CreateBorderDivs(wordDoc, settings, element.Elements())
            );
            cell.AddAnnotation(style);
            return cell;
        }

        private static object ProcessTableRow(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement element,
            decimal currentMarginLeft
        )
        {
            var style = new Dictionary<string, string>();
            var trHeight = (int?)element.Elements(W.trPr).Elements(W.trHeight).Attributes(W.val).FirstOrDefault();
            if (trHeight is not null)
                style.AddIfMissing("height", FormattableString.Invariant($"{(decimal)trHeight / 1440m:0.00}in"));
            var htmlRow = new XElement(
                Xhtml.tr,
                element.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft))
            );
            if (style.Any())
                htmlRow.AddAnnotation(style);
            return htmlRow;
        }

        private static bool HasStyleSeparator(XElement element)
        {
            return element is not null
                && element.Elements(W.pPr).Elements(W.rPr).Any(e => GetBoolProp(e, W.specVanish));
        }

        private static bool IsBidi(XElement element)
        {
            return element
                .Elements(W.pPr)
                .Elements(W.bidi)
                .Any(b => b.Attribute(W.val) is null || b.Attribute(W.val).ToBoolean() == true);
        }

        private static XName GetParagraphElementName(XElement element, WordprocessingDocument wordDoc)
        {
            var elementName = Xhtml.p;

            var styleId = (string)element.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
            if (styleId is null)
                return elementName;

            var style = GetStyle(styleId, wordDoc);
            if (style is null)
                return elementName;

            var outlineLevel = (int?)style.Elements(W.pPr).Elements(W.outlineLvl).Attributes(W.val).FirstOrDefault();
            if (outlineLevel is not null && outlineLevel <= 5)
            {
                elementName = Xhtml.xhtml + $"h{outlineLevel + 1}";
            }

            return elementName;
        }

        private static XElement GetStyle(string styleId, WordprocessingDocument wordDoc)
        {
            var stylesPart = wordDoc.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart is null)
                return null;

            var styles = stylesPart.GetXDocument().Root;
            return styles is not null
                ? styles.Elements(W.style).FirstOrDefault(s => (string)s.Attribute(W.styleId) == styleId)
                : null;
        }

        private static object CreateSectionDivs(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement element
        )
        {
            // note: when building a paging html converter, need to attend to new sections with page breaks here.
            // This code conflates adjacent sections if they have identical formatting, which is not an issue
            // for the non-paging transform.
            var groupedIntoDivs = element
                .Elements()
                .GroupAdjacent(e =>
                {
                    var sectAnnotation = e.Annotation<SectionAnnotation>();
                    return sectAnnotation is not null ? sectAnnotation.SectionElement.ToString() : "";
                });

            // note: when creating a paging html converter, need to pay attention to w:rtlGutter element.
            var divList = groupedIntoDivs.Select(g =>
            {
                var sectPr = g.First().Annotation<SectionAnnotation>();
                XElement bidi = null;
                if (sectPr is not null)
                {
                    bidi = sectPr
                        .SectionElement.Elements(W.bidi)
                        .FirstOrDefault(b => b.Attribute(W.val) is null || b.Attribute(W.val).ToBoolean() == true);
                }
                if (sectPr is null || bidi is null)
                {
                    var div = new XElement(Xhtml.div, CreateBorderDivs(wordDoc, settings, g));
                    return div;
                }
                else
                {
                    var div = new XElement(
                        Xhtml.div,
                        new XAttribute("dir", "rtl"),
                        CreateBorderDivs(wordDoc, settings, g)
                    );
                    return div;
                }
            });
            return divList;
        }

        private enum BorderType
        {
            Paragraph,
            Cell,
        };

        /*
         * Notes on line spacing
         *
         * the w:line and w:lineRule attributes control spacing between lines - including between lines within a paragraph
         *
         * If w:spacing w:lineRule="auto" then
         *   w:spacing w:line is a percentage where 240 == 100%
         *
         *   (line value / 240) * 100 = percentage of line
         *
         * If w:spacing w:lineRule="exact" or w:lineRule="atLeast" then
         *   w:spacing w:line is in twips
         *   1440 = exactly one inch from line to line
         *
         * Handle
         * - ind
         * - jc
         * - numPr
         * - pBdr
         * - shd
         * - spacing
         * - textAlignment
         *
         * Don't Handle (yet)
         * - adjustRightInd?
         * - autoSpaceDE
         * - autoSpaceDN
         * - bidi
         * - contextualSpacing
         * - divId
         * - framePr
         * - keepLines
         * - keepNext
         * - kinsoku
         * - mirrorIndents
         * - overflowPunct
         * - snapToGrid
         * - suppressAutoHyphens
         * - suppressLineNumbers
         * - suppressOverlap
         * - tabs
         * - textBoxTightWrap
         * - textDirection
         * - topLinePunct
         * - widowControl
         * - wordWrap
         *
         */

        private static object ConvertParagraph(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement paragraph,
            XName elementName,
            bool suppressTrailingWhiteSpace,
            decimal currentMarginLeft,
            bool isBidi
        )
        {
            var style = DefineParagraphStyle(
                paragraph,
                elementName,
                suppressTrailingWhiteSpace,
                currentMarginLeft,
                isBidi
            );
            var rtl = isBidi ? new XAttribute("dir", "rtl") : new XAttribute("dir", "ltr");
            var firstMark = isBidi ? new XEntity("#x200f") : null;

            // Analyze initial runs to see whether we have a tab, in which case we will render
            // a span with a defined width and ignore the tab rather than rendering the text
            // preceding the tab and the tab as a span with a computed width.
            var firstTabRun = paragraph.Elements(W.r).FirstOrDefault(run => run.Elements(W.tab).Any());
            var elementsPrecedingTab = firstTabRun is not null
                ? paragraph
                    .Elements(W.r)
                    .TakeWhile(e => e != firstTabRun)
                    .Where(e => e.Elements().Any(c => c.Attributes(PtOpenXml.TabWidth).Any()))
                    .ToList()
                : Enumerable.Empty<XElement>().ToList();

            // TODO: Revisit
            // For the time being, if a hyperlink field precedes the tab, we'll render it as before.
            var hyperlinkPrecedesTab = elementsPrecedingTab
                .Elements(W.r)
                .Elements(W.instrText)
                .Select(e => e.Value)
                .Any(value =>
                    value is not null && value.TrimStart().StartsWith("HYPERLINK", StringComparison.OrdinalIgnoreCase)
                );
            if (hyperlinkPrecedesTab)
            {
                var paraElement1 = new XElement(
                    elementName,
                    rtl,
                    firstMark,
                    ConvertContentThatCanContainFields(wordDoc, settings, paragraph.Elements())
                );
                paraElement1.AddAnnotation(style);
                return paraElement1;
            }

            var txElementsPrecedingTab = TransformElementsPrecedingTab(
                wordDoc,
                settings,
                elementsPrecedingTab,
                firstTabRun
            );
            var elementsSucceedingTab = firstTabRun is not null
                ? paragraph.Elements().SkipWhile(e => e != firstTabRun).Skip(1)
                : paragraph.Elements();
            var paraElement = new XElement(
                elementName,
                rtl,
                firstMark,
                txElementsPrecedingTab,
                ConvertContentThatCanContainFields(wordDoc, settings, elementsSucceedingTab)
            );
            paraElement.AddAnnotation(style);

            return paraElement;
        }

        private static List<object> TransformElementsPrecedingTab(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            List<XElement> elementsPrecedingTab,
            XElement firstTabRun
        )
        {
            var tabWidth = firstTabRun is not null
                ? (decimal?)firstTabRun.Elements(W.tab).Attributes(PtOpenXml.TabWidth).FirstOrDefault() ?? 0m
                : 0m;
            var precedingElementsWidth = elementsPrecedingTab
                .Elements()
                .Where(c => c.Attributes(PtOpenXml.TabWidth).Any())
                .Select(e => (decimal)e.Attribute(PtOpenXml.TabWidth))
                .Sum();
            var totalWidth = precedingElementsWidth + tabWidth;

            var txElementsPrecedingTab = elementsPrecedingTab
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .ToList();
            if (txElementsPrecedingTab.Count > 1)
            {
                var span = new XElement(Xhtml.span, txElementsPrecedingTab);
                var spanStyle = new Dictionary<string, string>
                {
                    { "display", "inline-block" },
                    { "text-indent", "0" },
                    // Use min-width so the span expands when text exceeds the tab stop width,
                    // preventing text overflow and overlap with subsequent content.
                    { "min-width", FormattableString.Invariant($"{totalWidth:0.000}in") },
                };
                span.AddAnnotation(spanStyle);

                // Replace the preceding elements with the wrapper span so that the
                // min-width styling takes effect on the returned content.
                txElementsPrecedingTab.Clear();
                txElementsPrecedingTab.Add(span);
            }
            else if (txElementsPrecedingTab.Count == 1)
            {
                var element = txElementsPrecedingTab.First() as XElement;
                if (element is not null)
                {
                    var spanStyle = element.Annotation<Dictionary<string, string>>();
                    spanStyle.AddIfMissing("display", "inline-block");
                    spanStyle.AddIfMissing("text-indent", "0");
                    spanStyle.AddIfMissing("min-width", FormattableString.Invariant($"{totalWidth:0.000}in"));
                }
            }
            return txElementsPrecedingTab;
        }

        private static Dictionary<string, string> DefineParagraphStyle(
            XElement paragraph,
            XName elementName,
            bool suppressTrailingWhiteSpace,
            decimal currentMarginLeft,
            bool isBidi
        )
        {
            var style = new Dictionary<string, string>();

            var styleName = (string)paragraph.Attribute(PtOpenXml.StyleName);
            if (styleName is not null)
                style.Add("PtStyleName", styleName);

            var pPr = paragraph.Element(W.pPr);
            if (pPr is null)
                return style;

            CreateStyleFromSpacing(style, pPr.Element(W.spacing), elementName, suppressTrailingWhiteSpace);
            CreateStyleFromInd(style, pPr.Element(W.ind), elementName, currentMarginLeft, isBidi);

            // todo need to handle
            // - both
            // - mediumKashida
            // - distribute
            // - numTab
            // - highKashida
            // - lowKashida
            // - thaiDistribute

            CreateStyleFromJc(style, pPr.Element(W.jc), isBidi);
            CreateStyleFromShd(style, pPr.Element(W.shd));

            if (GetBoolProp(pPr, W.pageBreakBefore))
                style.AddIfMissing("page-break-before", "always");

            // Pt.FontName
            var font = (string)paragraph.Attributes(PtOpenXml.FontName).FirstOrDefault();
            if (font is not null)
                CreateFontCssProperty(font, style);

            DefineFontSize(style, paragraph);
            DefineLineHeight(style, paragraph);

            // vertical text alignment as of December 2013 does not work in any major browsers.
            CreateStyleFromTextAlignment(style, pPr.Element(W.textAlignment));

            style.AddIfMissing("margin-top", "0");
            style.AddIfMissing("margin-left", "0");
            style.AddIfMissing("margin-right", "0");
            style.AddIfMissing("margin-bottom", ".001pt");

            return style;
        }

        private static void CreateStyleFromInd(
            Dictionary<string, string> style,
            XElement ind,
            XName elementName,
            decimal currentMarginLeft,
            bool isBidi
        )
        {
            if (ind is null)
                return;

            var left = (decimal?)ind.Attribute(W.left);
            if (left is not null && elementName != Xhtml.span)
            {
                var leftInInches = (decimal)left / 1440 - currentMarginLeft;
                style.AddIfMissing(
                    isBidi ? "margin-right" : "margin-left",
                    leftInInches > 0m ? FormattableString.Invariant($"{leftInInches:0.00}in") : "0"
                );
            }

            var right = (decimal?)ind.Attribute(W.right);
            if (right is not null)
            {
                var rightInInches = (decimal)right / 1440;
                style.AddIfMissing(
                    isBidi ? "margin-left" : "margin-right",
                    rightInInches > 0m ? FormattableString.Invariant($"{rightInInches:0.00}in") : "0"
                );
            }

            var firstLine = (decimal?)ind.Attribute(W.firstLine);
            if (firstLine is not null && elementName != Xhtml.span)
            {
                var firstLineInInches = (decimal)firstLine / 1440m;
                style.AddIfMissing("text-indent", FormattableString.Invariant($"{firstLineInInches:0.00}in"));
            }

            var hanging = (decimal?)ind.Attribute(W.hanging);
            if (hanging is not null && elementName != Xhtml.span)
            {
                var hangingInInches = (decimal)-hanging / 1440m;
                style.AddIfMissing("text-indent", FormattableString.Invariant($"{hangingInInches:0.00}in"));
            }
        }

        private static void CreateStyleFromJc(Dictionary<string, string> style, XElement jc, bool isBidi)
        {
            if (jc is not null)
            {
                var jcVal = (string)jc.Attributes(W.val).FirstOrDefault() ?? "left";
                if (jcVal == "left")
                    style.AddIfMissing("text-align", isBidi ? "right" : "left");
                else if (jcVal == "right")
                    style.AddIfMissing("text-align", isBidi ? "left" : "right");
                else if (jcVal == "center")
                    style.AddIfMissing("text-align", "center");
                else if (jcVal == "both")
                    style.AddIfMissing("text-align", "justify");
            }
        }

        private static void CreateStyleFromSpacing(
            Dictionary<string, string> style,
            XElement spacing,
            XName elementName,
            bool suppressTrailingWhiteSpace
        )
        {
            if (spacing is null)
                return;

            var spacingBefore = (decimal?)spacing.Attribute(W.before);
            if (spacingBefore is not null && elementName != Xhtml.span)
                style.AddIfMissing(
                    "margin-top",
                    spacingBefore > 0m ? FormattableString.Invariant($"{spacingBefore / 20.0m}pt") : "0"
                );

            var lineRule = (string)spacing.Attribute(W.lineRule);
            if (lineRule == "auto")
            {
                var line = (decimal)spacing.Attribute(W.line);
                if (line != 240m)
                {
                    var pct = (line / 240m) * 100m;
                    style.Add("line-height", FormattableString.Invariant($"{pct:0.0}%"));
                }
            }
            if (lineRule == "exact")
            {
                var line = (decimal)spacing.Attribute(W.line);
                var points = line / 20m;
                style.Add("line-height", FormattableString.Invariant($"{points:0.0}pt"));
            }
            if (lineRule == "atLeast")
            {
                var line = (decimal)spacing.Attribute(W.line);
                var points = line / 20m;
                if (points >= 14m)
                    style.Add("line-height", FormattableString.Invariant($"{points:0.0}pt"));
            }

            var spacingAfter = suppressTrailingWhiteSpace ? 0m : (decimal?)spacing.Attribute(W.after);
            if (spacingAfter is not null)
                style.AddIfMissing(
                    "margin-bottom",
                    spacingAfter > 0m ? FormattableString.Invariant($"{spacingAfter / 20.0m}pt") : "0"
                );
        }

        private static void CreateStyleFromTextAlignment(Dictionary<string, string> style, XElement textAlignment)
        {
            if (textAlignment is null)
                return;

            var verticalTextAlignment = (string)textAlignment.Attributes(W.val).FirstOrDefault();
            if (verticalTextAlignment is null or "auto")
                return;

            if (verticalTextAlignment == "top")
                style.AddIfMissing("vertical-align", "top");
            else if (verticalTextAlignment == "center")
                style.AddIfMissing("vertical-align", "middle");
            else if (verticalTextAlignment == "baseline")
                style.AddIfMissing("vertical-align", "baseline");
            else if (verticalTextAlignment == "bottom")
                style.AddIfMissing("vertical-align", "bottom");
        }

        private static void DefineFontSize(Dictionary<string, string> style, XElement paragraph)
        {
            var sz = paragraph
                .DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.r)
                .Select(WordprocessingMLUtil.GetFontSize)
                .Max();
            if (sz is not null)
                style.AddIfMissing("font-size", FormattableString.Invariant($"{sz / 2.0m}pt"));
        }

        private static void DefineLineHeight(Dictionary<string, string> style, XElement paragraph)
        {
            var allRunsAreUniDirectional = paragraph
                .DescendantsTrimmed(W.txbxContent)
                .Where(e => e.Name == W.r)
                .Select(run => (string)run.Attribute(PtOpenXml.LanguageType))
                .All(lt => lt != "bidi");
            if (allRunsAreUniDirectional)
                style.AddIfMissing("line-height", "108%");
        }

        /*
         * Handle:
         * - b
         * - bdr
         * - caps
         * - color
         * - dstrike
         * - highlight
         * - i
         * - position
         * - rFonts
         * - shd
         * - smallCaps
         * - spacing
         * - strike
         * - sz
         * - u
         * - vanish
         * - vertAlign
         *
         * Don't handle:
         * - em
         * - emboss
         * - fitText
         * - imprint
         * - kern
         * - outline
         * - shadow
         * - w
         *
         */

        private static object ConvertRun(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement run
        )
        {
            var rPr = run.Element(W.rPr);
            if (rPr is null)
                return run.Elements().Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m));

            // hide all content that contains the w:rPr/w:webHidden element
            if (rPr.Element(W.webHidden) is not null)
                return null;

            var style = DefineRunStyle(run);
            var convertedChildren = run.Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .Where(x => x is not null)
                .ToList();

            // If the run contains a single block-level <div> (e.g. a text box), return it directly
            // without any wrapping — a <span><div>…</div></span> would be invalid HTML.
            if (convertedChildren is [XElement singleDiv] && singleDiv.Name == Xhtml.div)
                return singleDiv;

            object content = convertedChildren;

            // Wrap content in h:sup or h:sub elements as necessary.
            if (rPr.Element(W.vertAlign) is not null)
            {
                XElement newContent = null;
                var vertAlignVal = (string)rPr.Elements(W.vertAlign).Attributes(W.val).FirstOrDefault();
                newContent = vertAlignVal switch
                {
                    "superscript" => new XElement(Xhtml.sup, content),
                    "subscript" => new XElement(Xhtml.sub, content),
                    _ => newContent,
                };
                if (newContent is not null && newContent.Nodes().Any())
                    content = newContent;
            }

            var langAttribute = GetLangAttribute(run);

            DetermineRunMarks(run, rPr, style, out var runStartMark, out var runEndMark);

            if (style.Any() || langAttribute is not null || runStartMark is not null)
            {
                style.AddIfMissing("margin", "0");
                style.AddIfMissing("padding", "0");
                var xe = new XElement(Xhtml.span, langAttribute, runStartMark, content, runEndMark);

                xe.AddAnnotation(style);
                content = xe;
            }
            return content;
        }

        [SuppressMessage("ReSharper", "FunctionComplexityOverflow")]
        private static Dictionary<string, string> DefineRunStyle(XElement run)
        {
            var style = new Dictionary<string, string>();

            var rPr = run.Elements(W.rPr).First();

            var styleName = (string)run.Attribute(PtOpenXml.StyleName);
            if (styleName is not null)
                style.Add("PtStyleName", styleName);

            // W.bdr
            if (
                rPr.Element(W.bdr) is not null
                && (string)rPr.Elements(W.bdr).Attributes(W.val).FirstOrDefault() != "none"
            )
            {
                style.AddIfMissing("border", "solid windowtext 1.0pt");
                style.AddIfMissing("padding", "0");
            }

            // W.color
            var color = (string)rPr.Elements(W.color).Attributes(W.val).FirstOrDefault();
            if (color is not null)
                CreateColorProperty("color", color, style);

            // W.highlight
            var highlight = (string)rPr.Elements(W.highlight).Attributes(W.val).FirstOrDefault();
            if (highlight is not null)
                CreateColorProperty("background", highlight, style);

            // W.shd
            var shade = (string)rPr.Elements(W.shd).Attributes(W.fill).FirstOrDefault();
            if (shade is not null)
                CreateColorProperty("background", shade, style);

            // Pt.FontName
            var sym = run.Element(W.sym);
            var font = sym is not null
                ? (string)sym.Attributes(W.font).FirstOrDefault()
                : (string)run.Attributes(PtOpenXml.FontName).FirstOrDefault();
            if (font is not null)
                CreateFontCssProperty(font, style);

            // W.sz
            var languageType = (string)run.Attribute(PtOpenXml.LanguageType);
            var sz = WordprocessingMLUtil.GetFontSize(languageType, rPr);
            if (sz is not null)
                style.AddIfMissing("font-size", FormattableString.Invariant($"{sz / 2.0m}pt"));

            // W.caps
            if (GetBoolProp(rPr, W.caps))
                style.AddIfMissing("text-transform", "uppercase");

            // W.smallCaps
            if (GetBoolProp(rPr, W.smallCaps))
                style.AddIfMissing("font-variant", "small-caps");

            // W.spacing
            var spacingInTwips = (decimal?)rPr.Elements(W.spacing).Attributes(W.val).FirstOrDefault();
            if (spacingInTwips is not null)
                style.AddIfMissing(
                    "letter-spacing",
                    spacingInTwips > 0m ? FormattableString.Invariant($"{spacingInTwips / 20}pt") : "0"
                );

            // W.position
            var position = (decimal?)rPr.Elements(W.position).Attributes(W.val).FirstOrDefault();
            if (position is not null)
            {
                style.AddIfMissing("position", "relative");
                style.AddIfMissing("top", FormattableString.Invariant($"{-(position / 2)}pt"));
            }

            // W.vanish
            if (GetBoolProp(rPr, W.vanish) && !GetBoolProp(rPr, W.specVanish))
                style.AddIfMissing("display", "none");

            // W.u
            if (rPr.Element(W.u) is not null && (string)rPr.Elements(W.u).Attributes(W.val).FirstOrDefault() != "none")
                style.AddIfMissing("text-decoration", "underline");

            // W.i
            style.AddIfMissing("font-style", GetBoolProp(rPr, W.i) ? "italic" : "normal");

            // W.b
            style.AddIfMissing("font-weight", GetBoolProp(rPr, W.b) ? "bold" : "normal");

            // W.strike
            if (GetBoolProp(rPr, W.strike) || GetBoolProp(rPr, W.dstrike))
                style.AddIfMissing("text-decoration", "line-through");

            return style;
        }

        private static void DetermineRunMarks(
            XElement run,
            XElement rPr,
            Dictionary<string, string> style,
            out XEntity runStartMark,
            out XEntity runEndMark
        )
        {
            runStartMark = null;
            runEndMark = null;

            // Only do the following for text runs.
            if (run.Element(W.t) is null)
                return;

            // Can't add directional marks if the font-family is a symbol/dingbat font —
            // these fonts use non-standard encodings where directional mark code points render as '?'.
            var addDirectionalMarks = true;
            if (style.TryGetValue("font-family", out var fontFamily))
            {
                var unquotedFontFamily = fontFamily;
                if (
                    unquotedFontFamily.Length >= 2
                    && (
                        (unquotedFontFamily[0] == '\'' && unquotedFontFamily[^1] == '\'')
                        || (unquotedFontFamily[0] == '"' && unquotedFontFamily[^1] == '"')
                    )
                )
                {
                    unquotedFontFamily = unquotedFontFamily[1..^1];
                }

                if (s_symbolFonts.Contains(unquotedFontFamily))
                    addDirectionalMarks = false;
            }
            if (!addDirectionalMarks)
                return;

            var isRtl = rPr.Element(W.rtl) is not null;
            if (isRtl)
            {
                runStartMark = new XEntity("#x200f"); // RLM
                runEndMark = new XEntity("#x200f"); // RLM
            }
            else
            {
                var paragraph = run.Ancestors(W.p).First();
                var paraIsBidi = paragraph
                    .Elements(W.pPr)
                    .Elements(W.bidi)
                    .Any(b => b.Attribute(W.val) is null || b.Attribute(W.val).ToBoolean() == true);

                if (paraIsBidi)
                {
                    runStartMark = new XEntity("#x200e"); // LRM
                    runEndMark = new XEntity("#x200e"); // LRM
                }
            }
        }

        private static XAttribute GetLangAttribute(XElement run)
        {
            const string defaultLanguage = "en-US"; // todo need to get defaultLanguage

            var rPr = run.Elements(W.rPr).First();
            var languageType = (string)run.Attribute(PtOpenXml.LanguageType);

            var lang = languageType switch
            {
                "western" => (string)rPr.Elements(W.lang).Attributes(W.val).FirstOrDefault(),
                "bidi" => (string)rPr.Elements(W.lang).Attributes(W.bidi).FirstOrDefault(),
                "eastAsia" => (string)rPr.Elements(W.lang).Attributes(W.eastAsia).FirstOrDefault(),
                _ => null,
            };
            if (lang is null)
                lang = defaultLanguage;

            return lang != defaultLanguage ? new XAttribute("lang", lang) : null;
        }

        private static void AdjustTableBorders(WordprocessingDocument wordDoc)
        {
            // Note: when implementing a paging version of the HTML transform, this needs to be done
            // for all content parts, not just the main document part.

            var xd = wordDoc.MainDocumentPart.GetXDocument();
            foreach (var tbl in xd.Descendants(W.tbl))
                AdjustTableBorders(tbl);
            wordDoc.MainDocumentPart.PutXDocument();
        }

        private static void AdjustTableBorders(XElement tbl)
        {
            var ta = tbl.Elements(W.tr)
                .Select(r =>
                    r.Elements(W.tc)
                        .SelectMany(c =>
                            Enumerable.Repeat(
                                c,
                                (int?)c.Elements(W.tcPr).Elements(W.gridSpan).Attributes(W.val).FirstOrDefault() ?? 1
                            )
                        )
                        .ToArray()
                )
                .ToArray();

            for (var y = 0; y < ta.Length; y++)
            {
                for (var x = 0; x < ta[y].Length; x++)
                {
                    var thisCell = ta[y][x];
                    FixTopBorder(ta, thisCell, x, y);
                    FixLeftBorder(ta, thisCell, x, y);
                    FixBottomBorder(ta, thisCell, x, y);
                    FixRightBorder(ta, thisCell, x, y);
                }
            }
        }

        private static void FixTopBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (y > 0)
            {
                var rowAbove = ta[y - 1];
                if (x < rowAbove.Length - 1)
                {
                    var cellAbove = ta[y - 1][x];
                    if (
                        cellAbove is not null
                        && thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() is not null
                        && cellAbove.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() is not null
                    )
                    {
                        ResolveCellBorder(
                            thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.top).FirstOrDefault(),
                            cellAbove.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.bottom).FirstOrDefault()
                        );
                    }
                }
            }
        }

        private static void FixLeftBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (x > 0)
            {
                var cellLeft = ta[y][x - 1];
                if (
                    cellLeft is not null
                    && thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() is not null
                    && cellLeft.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() is not null
                )
                {
                    ResolveCellBorder(
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.left).FirstOrDefault(),
                        cellLeft.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.right).FirstOrDefault()
                    );
                }
            }
        }

        private static void FixBottomBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (y < ta.Length - 1)
            {
                var rowBelow = ta[y + 1];
                if (x < rowBelow.Length - 1)
                {
                    var cellBelow = ta[y + 1][x];
                    if (
                        cellBelow is not null
                        && thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() is not null
                        && cellBelow.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() is not null
                    )
                    {
                        ResolveCellBorder(
                            thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.bottom).FirstOrDefault(),
                            cellBelow.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.top).FirstOrDefault()
                        );
                    }
                }
            }
        }

        private static void FixRightBorder(XElement[][] ta, XElement thisCell, int x, int y)
        {
            if (x < ta[y].Length - 1)
            {
                var cellRight = ta[y][x + 1];
                if (
                    cellRight is not null
                    && thisCell.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() is not null
                    && cellRight.Elements(W.tcPr).Elements(W.tcBorders).FirstOrDefault() is not null
                )
                {
                    ResolveCellBorder(
                        thisCell.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.right).FirstOrDefault(),
                        cellRight.Elements(W.tcPr).Elements(W.tcBorders).Elements(W.left).FirstOrDefault()
                    );
                }
            }
        }

        private static readonly Dictionary<string, int> BorderTypePriority = new()
        {
            { "single", 1 },
            { "thick", 2 },
            { "double", 3 },
            { "dotted", 4 },
        };

        private static readonly Dictionary<string, int> BorderNumber = new()
        {
            { "single", 1 },
            { "thick", 2 },
            { "double", 3 },
            { "dotted", 4 },
            { "dashed", 5 },
            { "dotDash", 5 },
            { "dotDotDash", 5 },
            { "triple", 6 },
            { "thinThickSmallGap", 6 },
            { "thickThinSmallGap", 6 },
            { "thinThickThinSmallGap", 6 },
            { "thinThickMediumGap", 6 },
            { "thickThinMediumGap", 6 },
            { "thinThickThinMediumGap", 6 },
            { "thinThickLargeGap", 6 },
            { "thickThinLargeGap", 6 },
            { "thinThickThinLargeGap", 6 },
            { "wave", 7 },
            { "doubleWave", 7 },
            { "dashSmallGap", 5 },
            { "dashDotStroked", 5 },
            { "threeDEmboss", 7 },
            { "threeDEngrave", 7 },
            { "outset", 7 },
            { "inset", 7 },
        };

        private static void ResolveCellBorder(XElement border1, XElement border2)
        {
            if (border1 is null || border2 is null)
                return;
            if ((string)border1.Attribute(W.val) == "nil" || (string)border2.Attribute(W.val) == "nil")
                return;
            if ((string)border1.Attribute(W.sz) == "nil" || (string)border2.Attribute(W.sz) == "nil")
                return;

            var border1Val = (string)border1.Attribute(W.val);
            var border1Weight = 1;
            if (BorderNumber.TryGetValue(border1Val, out var border1WeightValue))
                border1Weight = border1WeightValue;

            var border2Val = (string)border2.Attribute(W.val);
            var border2Weight = 1;
            if (BorderNumber.TryGetValue(border2Val, out var border2WeightValue))
                border2Weight = border2WeightValue;

            if (border1Weight != border2Weight)
            {
                if (border1Weight < border2Weight)
                    BorderOverride(border2, border1);
                else
                    BorderOverride(border1, border2);
            }

            if ((decimal)border1.Attribute(W.sz) > (decimal)border2.Attribute(W.sz))
            {
                BorderOverride(border1, border2);
                return;
            }

            if ((decimal)border1.Attribute(W.sz) < (decimal)border2.Attribute(W.sz))
            {
                BorderOverride(border2, border1);
                return;
            }

            var border1Type = (string)border1.Attribute(W.val);
            var border2Type = (string)border2.Attribute(W.val);
            if (
                BorderTypePriority.TryGetValue(border1Type, out var border1Pri)
                && BorderTypePriority.TryGetValue(border2Type, out var border2Pri)
            )
            {
                if (border1Pri < border2Pri)
                {
                    BorderOverride(border2, border1);
                    return;
                }
                if (border2Pri < border1Pri)
                {
                    BorderOverride(border1, border2);
                    return;
                }
            }

            var color1Str = (string)border1.Attribute(W.color);
            if (color1Str == "auto")
                color1Str = "000000";
            var color2Str = (string)border2.Attribute(W.color);
            if (color2Str == "auto")
                color2Str = "000000";
            if (color1Str is not null && color2Str is not null && color1Str != color2Str)
            {
                try
                {
                    var color1 = Convert.ToInt32(color1Str, 16);
                    var color2 = Convert.ToInt32(color2Str, 16);
                    if (color1 < color2)
                    {
                        BorderOverride(border1, border2);
                        return;
                    }
                    if (color2 < color1)
                    {
                        BorderOverride(border2, border1);
                    }
                }
                // if the above throws ArgumentException, FormatException, or OverflowException, then abort
                catch (Exception)
                {
                    // Ignore
                }
            }
        }

        private static void BorderOverride(XElement fromBorder, XElement toBorder)
        {
            toBorder.Attribute(W.val).Value = fromBorder.Attribute(W.val).Value;
            if (fromBorder.Attribute(W.color) is not null)
                toBorder.SetAttributeValue(W.color, fromBorder.Attribute(W.color).Value);
            if (fromBorder.Attribute(W.sz) is not null)
                toBorder.SetAttributeValue(W.sz, fromBorder.Attribute(W.sz).Value);
            if (fromBorder.Attribute(W.themeColor) is not null)
                toBorder.SetAttributeValue(W.themeColor, fromBorder.Attribute(W.themeColor).Value);
            if (fromBorder.Attribute(W.themeTint) is not null)
                toBorder.SetAttributeValue(W.themeTint, fromBorder.Attribute(W.themeTint).Value);
        }

        private static void CalculateSpanWidthForTabs(WordprocessingDocument wordDoc)
        {
            // Note: when implementing a paging version of the HTML transform, this needs to be done
            // for all content parts, not just the main document part.

            // w:defaultTabStop in settings
            var sxd = wordDoc.MainDocumentPart.DocumentSettingsPart.GetXDocument();
            var defaultTabStop = (int?)sxd.Descendants(W.defaultTabStop).Attributes(W.val).FirstOrDefault() ?? 720;

            var pxd = wordDoc.MainDocumentPart.GetXDocument();
            var root = pxd.Root;
            if (root is null)
                return;

            var newRoot = (XElement)CalculateSpanWidthTransform(root, defaultTabStop);
            root.ReplaceWith(newRoot);
            wordDoc.MainDocumentPart.PutXDocument();
        }

        // TODO: Refactor. This method is way too long.
        [SuppressMessage("ReSharper", "FunctionComplexityOverflow")]
        private static object CalculateSpanWidthTransform(XNode node, int defaultTabStop)
        {
            if (node is not XElement element)
                return node;

            // if it is not a paragraph or if there are no tabs in the paragraph,
            // then no need to continue processing.
            if (
                element.Name != W.p
                || !element.DescendantsTrimmed(W.txbxContent).Where(d => d.Name == W.r).Elements(W.tab).Any()
            )
            {
                // TODO: Revisit. Can we just return the node if it is a paragraph that does not have any tab?
                return new XElement(
                    element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => CalculateSpanWidthTransform(n, defaultTabStop))
                );
            }

            var clonedPara = new XElement(element);

            var leftInTwips = 0;
            var firstInTwips = 0;

            var ind = clonedPara.Elements(W.pPr).Elements(W.ind).FirstOrDefault();
            if (ind is not null)
            {
                // todo need to handle start and end attributes

                var left = (int?)ind.Attribute(W.left);
                if (left is not null)
                    leftInTwips = (int)left;

                var firstLine = 0;
                var firstLineAtt = (int?)ind.Attribute(W.firstLine);
                if (firstLineAtt is not null)
                    firstLine = (int)firstLineAtt;

                var hangingAtt = (int?)ind.Attribute(W.hanging);
                if (hangingAtt is not null)
                    firstLine = -(int)hangingAtt;

                firstInTwips = leftInTwips + firstLine;
            }

            // calculate the tab stops, in twips
            var tabs = clonedPara.Elements(W.pPr).Elements(W.tabs).FirstOrDefault();

            if (tabs is null)
            {
                if (leftInTwips == 0)
                {
                    tabs = new XElement(
                        W.tabs,
                        Enumerable
                            .Range(1, 100)
                            .Select(r => new XElement(
                                W.tab,
                                new XAttribute(W.val, "left"),
                                new XAttribute(W.pos, r * defaultTabStop)
                            ))
                    );
                }
                else
                {
                    tabs = new XElement(
                        W.tabs,
                        new XElement(W.tab, new XAttribute(W.val, "left"), new XAttribute(W.pos, leftInTwips))
                    );
                    tabs = AddDefaultTabsAfterLastTab(tabs, defaultTabStop);
                }
            }
            else
            {
                if (leftInTwips != 0)
                {
                    tabs.Add(new XElement(W.tab, new XAttribute(W.val, "left"), new XAttribute(W.pos, leftInTwips)));
                }
                tabs = AddDefaultTabsAfterLastTab(tabs, defaultTabStop);
            }

            var twipCounter = firstInTwips;
            var contentToMeasure = element
                .DescendantsTrimmed(z => z.Name == W.txbxContent || z.Name == W.pPr || z.Name == W.rPr)
                .ToArray();
            var currentElementIdx = 0;
            while (true)
            {
                var currentElement = contentToMeasure[currentElementIdx];

                if (currentElement.Name == W.br)
                {
                    twipCounter = leftInTwips;

                    currentElement.Add(
                        new XAttribute(PtOpenXml.TabWidth, FormattableString.Invariant($"{firstInTwips / 1440m:0.000}"))
                    );

                    currentElementIdx++;
                    if (currentElementIdx >= contentToMeasure.Length)
                        break; // we're done
                }

                if (currentElement.Name == W.tab)
                {
                    var runContainingTabToReplace = currentElement.Parent;
                    var fontNameAtt =
                        runContainingTabToReplace.Attribute(PtOpenXml.pt + "FontName")
                        ?? runContainingTabToReplace.Ancestors(W.p).First().Attribute(PtOpenXml.pt + "FontName");

                    var testAmount = twipCounter;

                    var tabAfterText = tabs.Elements(W.tab).FirstOrDefault(t => (int)t.Attribute(W.pos) > testAmount);

                    if (tabAfterText is null)
                    {
                        // something has gone wrong, so put 1/2 inch in
                        if (currentElement.Attribute(PtOpenXml.TabWidth) is null)
                            currentElement.Add(new XAttribute(PtOpenXml.TabWidth, 720m));
                        break;
                    }

                    var tabVal = (string)tabAfterText.Attribute(W.val);
                    if (tabVal is "right" or "end")
                    {
                        var textAfterElements = contentToMeasure.Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements
                            .TakeWhile(z => z.Name != W.tab && z.Name != W.br && z.Name != W.cr)
                            .ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        var dummyRun2 = new XElement(
                            W.r,
                            fontNameAtt,
                            runContainingTabToReplace.Elements(W.rPr),
                            new XElement(W.t, textAfterTab)
                        );

                        var widthOfTextAfterTab = WordprocessingMLUtil.CalcWidthOfRunInTwips(dummyRun2);
                        var delta2 = (int)tabAfterText.Attribute(W.pos) - widthOfTextAfterTab - twipCounter;
                        if (delta2 < 0)
                            delta2 = 0;
                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth, FormattableString.Invariant($"{delta2 / 1440m:0.000}")),
                            GetLeader(tabAfterText)
                        );
                        twipCounter = Math.Max((int)tabAfterText.Attribute(W.pos), twipCounter + widthOfTextAfterTab);

                        var lastElement = textElementsToMeasure.LastOrDefault();
                        if (lastElement is null)
                            break; // we're done

                        currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                        if (currentElementIdx >= contentToMeasure.Length)
                            break; // we're done

                        continue;
                    }
                    if (tabVal == "decimal")
                    {
                        var textAfterElements = contentToMeasure.Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements
                            .TakeWhile(z => z.Name != W.tab && z.Name != W.br && z.Name != W.cr)
                            .ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        if (textAfterTab.Contains("."))
                        {
                            var mantissa = textAfterTab.Split('.')[0];

                            var dummyRun4 = new XElement(
                                W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, mantissa)
                            );

                            var widthOfMantissa = WordprocessingMLUtil.CalcWidthOfRunInTwips(dummyRun4);
                            var delta2 = (int)tabAfterText.Attribute(W.pos) - widthOfMantissa - twipCounter;
                            if (delta2 < 0)
                                delta2 = 0;
                            currentElement.Add(
                                new XAttribute(
                                    PtOpenXml.TabWidth,
                                    FormattableString.Invariant($"{delta2 / 1440m:0.000}")
                                ),
                                GetLeader(tabAfterText)
                            );

                            var decims = textAfterTab.Substring(textAfterTab.IndexOf('.'));
                            dummyRun4 = new XElement(
                                W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, decims)
                            );

                            var widthOfDecims = WordprocessingMLUtil.CalcWidthOfRunInTwips(dummyRun4);
                            twipCounter = Math.Max(
                                (int)tabAfterText.Attribute(W.pos) + widthOfDecims,
                                twipCounter + widthOfMantissa + widthOfDecims
                            );

                            var lastElement = textElementsToMeasure.LastOrDefault();
                            if (lastElement is null)
                                break; // we're done

                            currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                            if (currentElementIdx >= contentToMeasure.Length)
                                break; // we're done

                            continue;
                        }
                        else
                        {
                            var dummyRun2 = new XElement(
                                W.r,
                                fontNameAtt,
                                runContainingTabToReplace.Elements(W.rPr),
                                new XElement(W.t, textAfterTab)
                            );

                            var widthOfTextAfterTab = WordprocessingMLUtil.CalcWidthOfRunInTwips(dummyRun2);
                            var delta2 = (int)tabAfterText.Attribute(W.pos) - widthOfTextAfterTab - twipCounter;
                            if (delta2 < 0)
                                delta2 = 0;
                            currentElement.Add(
                                new XAttribute(
                                    PtOpenXml.TabWidth,
                                    FormattableString.Invariant($"{delta2 / 1440m:0.000}")
                                ),
                                GetLeader(tabAfterText)
                            );
                            twipCounter = Math.Max(
                                (int)tabAfterText.Attribute(W.pos),
                                twipCounter + widthOfTextAfterTab
                            );

                            var lastElement = textElementsToMeasure.LastOrDefault();
                            if (lastElement is null)
                                break; // we're done

                            currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                            if (currentElementIdx >= contentToMeasure.Length)
                                break; // we're done

                            continue;
                        }
                    }
                    if ((string)tabAfterText.Attribute(W.val) == "center")
                    {
                        var textAfterElements = contentToMeasure.Skip(currentElementIdx + 1);

                        // take all the content until another tab, br, or cr
                        var textElementsToMeasure = textAfterElements
                            .TakeWhile(z => z.Name != W.tab && z.Name != W.br && z.Name != W.cr)
                            .ToList();

                        var textAfterTab = textElementsToMeasure
                            .Where(z => z.Name == W.t)
                            .Select(t => (string)t)
                            .StringConcatenate();

                        var dummyRun4 = new XElement(
                            W.r,
                            fontNameAtt,
                            runContainingTabToReplace.Elements(W.rPr),
                            new XElement(W.t, textAfterTab)
                        );

                        var widthOfText = WordprocessingMLUtil.CalcWidthOfRunInTwips(dummyRun4);
                        var delta2 = (int)tabAfterText.Attribute(W.pos) - (widthOfText / 2) - twipCounter;
                        if (delta2 < 0)
                            delta2 = 0;
                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth, FormattableString.Invariant($"{delta2 / 1440m:0.000}")),
                            GetLeader(tabAfterText)
                        );
                        twipCounter = Math.Max(
                            (int)tabAfterText.Attribute(W.pos) + widthOfText / 2,
                            twipCounter + widthOfText
                        );

                        var lastElement = textElementsToMeasure.LastOrDefault();
                        if (lastElement is null)
                            break; // we're done

                        currentElementIdx = Array.IndexOf(contentToMeasure, lastElement) + 1;
                        if (currentElementIdx >= contentToMeasure.Length)
                            break; // we're done

                        continue;
                    }
                    if (tabVal is "left" or "start" or "num")
                    {
                        var delta = (int)tabAfterText.Attribute(W.pos) - twipCounter;
                        currentElement.Add(
                            new XAttribute(PtOpenXml.TabWidth, FormattableString.Invariant($"{delta / 1440m:0.000}")),
                            GetLeader(tabAfterText)
                        );
                        twipCounter = (int)tabAfterText.Attribute(W.pos);

                        currentElementIdx++;
                        if (currentElementIdx >= contentToMeasure.Length)
                            break; // we're done

                        continue;
                    }
                }

                if (currentElement.Name == W.t)
                {
                    // TODO: Revisit. This is a quick fix because it doesn't work on Azure.
                    // Given the changes we've made elsewhere, though, this is not required
                    // for the first tab at least. We could also enhance that other change
                    // to deal with all tabs.
                    //var runContainingTabToReplace = currentElement.Parent;
                    //var paragraphForRun = runContainingTabToReplace.Ancestors(W.p).First();
                    //var fontNameAtt = runContainingTabToReplace.Attribute(PtOpenXml.FontName) ??
                    //                  paragraphForRun.Attribute(PtOpenXml.FontName);
                    //var languageTypeAtt = runContainingTabToReplace.Attribute(PtOpenXml.LanguageType) ??
                    //                      paragraphForRun.Attribute(PtOpenXml.LanguageType);

                    //var dummyRun3 = new XElement(W.r, fontNameAtt, languageTypeAtt,
                    //    runContainingTabToReplace.Elements(W.rPr),
                    //    currentElement);
                    //var widthOfText = CalcWidthOfRunInTwips(dummyRun3);
                    const int widthOfText = 0;
                    currentElement.Add(
                        new XAttribute(PtOpenXml.TabWidth, FormattableString.Invariant($"{widthOfText / 1440m:0.000}"))
                    );
                    twipCounter += widthOfText;

                    currentElementIdx++;
                    if (currentElementIdx >= contentToMeasure.Length)
                        break; // we're done

                    continue;
                }

                currentElementIdx++;
                if (currentElementIdx >= contentToMeasure.Length)
                    break; // we're done
            }

            return new XElement(
                element.Name,
                element.Attributes(),
                element.Nodes().Select(n => CalculateSpanWidthTransform(n, defaultTabStop))
            );
        }

        private static XAttribute GetLeader(XElement tabAfterText)
        {
            var leader = (string)tabAfterText.Attribute(W.leader);
            if (leader is null)
                return null;
            return new XAttribute(PtOpenXml.Leader, leader);
        }

        private static XElement AddDefaultTabsAfterLastTab(XElement tabs, int defaultTabStop)
        {
            var lastTabElement = tabs.Elements(W.tab)
                .Where(t => (string)t.Attribute(W.val) != "clear" && (string)t.Attribute(W.val) != "bar")
                .OrderBy(t => (int)t.Attribute(W.pos))
                .LastOrDefault();
            if (lastTabElement is not null)
            {
                if (defaultTabStop == 0)
                    defaultTabStop = 720;
                var rangeStart = (int)lastTabElement.Attribute(W.pos) / defaultTabStop + 1;
                var tempTabs = new XElement(
                    W.tabs,
                    tabs.Elements()
                        .Where(t => (string)t.Attribute(W.val) != "clear" && (string)t.Attribute(W.val) != "bar"),
                    Enumerable
                        .Range(rangeStart, 100)
                        .Select(r => new XElement(
                            W.tab,
                            new XAttribute(W.val, "left"),
                            new XAttribute(W.pos, r * defaultTabStop)
                        ))
                );
                tempTabs = new XElement(W.tabs, tempTabs.Elements().OrderBy(t => (int)t.Attribute(W.pos)));
                return tempTabs;
            }
            else
            {
                tabs = new XElement(
                    W.tabs,
                    Enumerable
                        .Range(1, 100)
                        .Select(r => new XElement(
                            W.tab,
                            new XAttribute(W.val, "left"),
                            new XAttribute(W.pos, r * defaultTabStop)
                        ))
                );
            }
            return tabs;
        }

        private static void InsertAppropriateNonbreakingSpaces(WordprocessingDocument wordDoc)
        {
            foreach (var part in wordDoc.ContentParts())
            {
                var pxd = part.GetXDocument();
                var root = pxd.Root;
                if (root is null)
                    return;

                var newRoot = (XElement)InsertAppropriateNonbreakingSpacesTransform(root);
                root.ReplaceWith(newRoot);
                part.PutXDocument();
            }
        }

        // Non-breaking spaces are not required if we use appropriate CSS, i.e., "white-space: pre-wrap;".
        // We only need to make sure that empty w:p elements are translated into non-empty h:p elements,
        // because empty h:p elements would be ignored by browsers.
        // Further, in addition to not being required, non-breaking spaces would change the layout behavior
        // of spans having consecutive spaces. Therefore, avoiding non-breaking spaces has the additional
        // benefit of leading to a more faithful representation of the Word document in HTML.
        private static object InsertAppropriateNonbreakingSpacesTransform(XNode node)
        {
            if (node is XElement element)
            {
                // child content of run to look for
                // W.br
                // W.cr
                // W.dayLong
                // W.dayShort
                // W.drawing
                // W.monthLong
                // W.monthShort
                // W.noBreakHyphen
                // W.object
                // W.pgNum
                // W.pTab
                // W.separator
                // W.softHyphen
                // W.sym
                // W.t
                // W.tab
                // W.yearLong
                // W.yearShort
                if (element.Name == W.p)
                {
                    // Translate empty paragraphs to paragraphs having one run with
                    // a normal space. A non-breaking space, i.e., \x00A0, is not
                    // required if we use appropriate CSS.
                    var hasContent = element
                        .Elements()
                        .Where(e => e.Name != W.pPr)
                        .DescendantsAndSelf()
                        .Any(e =>
                            e.Name == W.dayLong
                            || e.Name == W.dayShort
                            || e.Name == W.drawing
                            || e.Name == W.monthLong
                            || e.Name == W.monthShort
                            || e.Name == W.noBreakHyphen
                            || e.Name == W._object
                            || e.Name == W.pgNum
                            || e.Name == W.ptab
                            || e.Name == W.separator
                            || e.Name == W.softHyphen
                            || e.Name == W.sym
                            || e.Name == W.t
                            || e.Name == W.tab
                            || e.Name == W.yearLong
                            || e.Name == W.yearShort
                        );
                    if (hasContent == false)
                        return new XElement(
                            element.Name,
                            element.Attributes(),
                            element.Nodes().Select(n => InsertAppropriateNonbreakingSpacesTransform(n)),
                            new XElement(W.r, element.Elements(W.pPr).Elements(W.rPr), new XElement(W.t, " "))
                        );
                }

                return new XElement(
                    element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => InsertAppropriateNonbreakingSpacesTransform(n))
                );
            }
            return node;
        }

        private class SectionAnnotation
        {
            public XElement SectionElement;
        }

        private static void AnnotateForSections(WordprocessingDocument wordDoc)
        {
            var xd = wordDoc.MainDocumentPart.GetXDocument();

            var document = xd.Root;
            if (document is null)
                return;

            var body = document.Element(W.body);
            if (body is null)
                return;

            // move last sectPr into last paragraph
            var lastSectPr = body.Elements(W.sectPr).LastOrDefault();
            if (lastSectPr is not null)
            {
                // if the last thing in the document is a table, Word will always insert a paragraph following that.
                var lastPara = body.DescendantsTrimmed(W.txbxContent).LastOrDefault(p => p.Name == W.p);

                if (lastPara is not null)
                {
                    var lastParaProps = lastPara.Element(W.pPr);
                    if (lastParaProps is not null)
                        lastParaProps.Add(lastSectPr);
                    else
                        lastPara.Add(new XElement(W.pPr, lastSectPr));

                    lastSectPr.Remove();
                }
            }

            var reverseDescendants = xd.Descendants().Reverse().ToList();
            var currentSection = InitializeSectionAnnotation(reverseDescendants);

            foreach (var d in reverseDescendants)
            {
                if (d.Name == W.sectPr)
                {
                    if (d.Attribute(XNamespace.Xmlns + "w") is null)
                        d.Add(new XAttribute(XNamespace.Xmlns + "w", W.w));

                    currentSection = new SectionAnnotation() { SectionElement = d };
                }
                else
                    d.AddAnnotation(currentSection);
            }
        }

        private static SectionAnnotation InitializeSectionAnnotation(IEnumerable<XElement> reverseDescendants)
        {
            var currentSection = new SectionAnnotation()
            {
                SectionElement = reverseDescendants.FirstOrDefault(e => e.Name == W.sectPr),
            };
            if (
                currentSection.SectionElement is not null
                && currentSection.SectionElement.Attribute(XNamespace.Xmlns + "w") is null
            )
                currentSection.SectionElement.Add(new XAttribute(XNamespace.Xmlns + "w", W.w));

            // todo what should the default section props be?
            if (currentSection.SectionElement is null)
                currentSection = new SectionAnnotation()
                {
                    SectionElement = new XElement(
                        W.sectPr,
                        new XAttribute(XNamespace.Xmlns + "w", W.w),
                        new XElement(W.pgSz, new XAttribute(W._w, 12240), new XAttribute(W.h, 15840)),
                        new XElement(
                            W.pgMar,
                            new XAttribute(W.top, 1440),
                            new XAttribute(W.right, 1440),
                            new XAttribute(W.bottom, 1440),
                            new XAttribute(W.left, 1440),
                            new XAttribute(W.header, 720),
                            new XAttribute(W.footer, 720),
                            new XAttribute(W.gutter, 0)
                        ),
                        new XElement(W.cols, new XAttribute(W.space, 720)),
                        new XElement(W.docGrid, new XAttribute(W.linePitch, 360))
                    ),
                };

            return currentSection;
        }

        private static object CreateBorderDivs(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            IEnumerable<XElement> elements
        )
        {
            return elements
                .GroupAdjacent(e =>
                {
                    var pBdr = e.Elements(W.pPr).Elements(W.pBdr).FirstOrDefault();
                    if (pBdr is not null)
                    {
                        var indStr = string.Empty;
                        var ind = e.Elements(W.pPr).Elements(W.ind).FirstOrDefault();
                        if (ind is not null)
                            indStr = ind.ToString(SaveOptions.DisableFormatting);
                        return pBdr.ToString(SaveOptions.DisableFormatting) + indStr;
                    }
                    return e.Name == W.tbl ? "table" : string.Empty;
                })
                .Select(g =>
                {
                    if (g.Key == string.Empty)
                    {
                        return (object)GroupAndVerticallySpaceNumberedParagraphs(wordDoc, settings, g, 0m);
                    }
                    if (g.Key == "table")
                    {
                        return g.Select(gc => ConvertToHtmlTransform(wordDoc, settings, gc, false, 0));
                    }
                    var pPr = g.First().Elements(W.pPr).First();
                    var pBdr = pPr.Element(W.pBdr);
                    var style = new Dictionary<string, string>();
                    GenerateBorderStyle(pBdr, W.top, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.right, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.bottom, style, BorderType.Paragraph);
                    GenerateBorderStyle(pBdr, W.left, style, BorderType.Paragraph);

                    var currentMarginLeft = 0m;
                    var ind = pPr.Element(W.ind);
                    if (ind is not null)
                    {
                        var leftInInches = (decimal?)ind.Attribute(W.left) / 1440m ?? 0;
                        var hangingInInches = -(decimal?)ind.Attribute(W.hanging) / 1440m ?? 0;
                        currentMarginLeft = leftInInches + hangingInInches;

                        style.AddIfMissing(
                            "margin-left",
                            currentMarginLeft > 0m ? FormattableString.Invariant($"{currentMarginLeft:0.00}in") : "0"
                        );
                    }

                    var div = new XElement(
                        Xhtml.div,
                        GroupAndVerticallySpaceNumberedParagraphs(wordDoc, settings, g, currentMarginLeft)
                    );
                    div.AddAnnotation(style);
                    return div;
                })
                .ToList();
        }

        private static IEnumerable<object> GroupAndVerticallySpaceNumberedParagraphs(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            IEnumerable<XElement> elements,
            decimal currentMarginLeft
        )
        {
            var grouped = elements
                .GroupAdjacent(e =>
                {
                    var abstractNumId = (string)e.Attribute(PtOpenXml.pt + "AbstractNumId");
                    if (abstractNumId is not null)
                        return "num:" + abstractNumId;
                    var contextualSpacing = e.Elements(W.pPr).Elements(W.contextualSpacing).FirstOrDefault();
                    if (contextualSpacing is not null)
                    {
                        var styleName = (string)e.Elements(W.pPr).Elements(W.pStyle).Attributes(W.val).FirstOrDefault();
                        if (styleName is null)
                            return "";
                        return "sty:" + styleName;
                    }
                    return "";
                })
                .ToList();
            var newContent = grouped.Select(g =>
            {
                if (g.Key == "")
                    return g.Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, currentMarginLeft));
                var last = g.Count() - 1;
                return g.Select((e, i) => ConvertToHtmlTransform(wordDoc, settings, e, i != last, currentMarginLeft));
            });
            return newContent;
        }

        private class BorderMappingInfo
        {
            public string CssName;
            public decimal CssSize;
        }

        private static readonly Dictionary<string, BorderMappingInfo> BorderStyleMap = new()
        {
            {
                "single",
                new BorderMappingInfo() { CssName = "solid", CssSize = 1.0m }
            },
            {
                "dotted",
                new BorderMappingInfo() { CssName = "dotted", CssSize = 1.0m }
            },
            {
                "dashSmallGap",
                new BorderMappingInfo() { CssName = "dashed", CssSize = 1.0m }
            },
            {
                "dashed",
                new BorderMappingInfo() { CssName = "dashed", CssSize = 1.0m }
            },
            {
                "dotDash",
                new BorderMappingInfo() { CssName = "dashed", CssSize = 1.0m }
            },
            {
                "dotDotDash",
                new BorderMappingInfo() { CssName = "dashed", CssSize = 1.0m }
            },
            {
                "double",
                new BorderMappingInfo() { CssName = "double", CssSize = 2.5m }
            },
            {
                "triple",
                new BorderMappingInfo() { CssName = "double", CssSize = 2.5m }
            },
            {
                "thinThickSmallGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 4.5m }
            },
            {
                "thickThinSmallGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 4.5m }
            },
            {
                "thinThickThinSmallGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 6.0m }
            },
            {
                "thickThinMediumGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 6.0m }
            },
            {
                "thinThickMediumGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 6.0m }
            },
            {
                "thinThickThinMediumGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 9.0m }
            },
            {
                "thinThickLargeGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 5.25m }
            },
            {
                "thickThinLargeGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 5.25m }
            },
            {
                "thinThickThinLargeGap",
                new BorderMappingInfo() { CssName = "double", CssSize = 9.0m }
            },
            {
                "wave",
                new BorderMappingInfo() { CssName = "solid", CssSize = 3.0m }
            },
            {
                "doubleWave",
                new BorderMappingInfo() { CssName = "double", CssSize = 5.25m }
            },
            {
                "dashDotStroked",
                new BorderMappingInfo() { CssName = "solid", CssSize = 3.0m }
            },
            {
                "threeDEmboss",
                new BorderMappingInfo() { CssName = "ridge", CssSize = 6.0m }
            },
            {
                "threeDEngrave",
                new BorderMappingInfo() { CssName = "groove", CssSize = 6.0m }
            },
            {
                "outset",
                new BorderMappingInfo() { CssName = "outset", CssSize = 4.5m }
            },
            {
                "inset",
                new BorderMappingInfo() { CssName = "inset", CssSize = 4.5m }
            },
        };

        private static void GenerateBorderStyle(
            XElement pBdr,
            XName sideXName,
            Dictionary<string, string> style,
            BorderType borderType
        )
        {
            string whichSide;
            if (sideXName == W.top)
                whichSide = "top";
            else if (sideXName == W.right)
                whichSide = "right";
            else if (sideXName == W.bottom)
                whichSide = "bottom";
            else
                whichSide = "left";
            if (pBdr is null)
            {
                style.Add("border-" + whichSide, "none");
                if (borderType == BorderType.Cell && whichSide is "left" or "right")
                    style.Add("padding-" + whichSide, "5.4pt");
                return;
            }

            var side = pBdr.Element(sideXName);
            if (side is null)
            {
                style.Add("border-" + whichSide, "none");
                if (borderType == BorderType.Cell && whichSide is "left" or "right")
                    style.Add("padding-" + whichSide, "5.4pt");
                return;
            }
            var type = (string)side.Attribute(W.val);
            if (type is "nil" or "none")
            {
                style.Add("border-" + whichSide + "-style", "none");

                var space = (decimal?)side.Attribute(W.space) ?? 0;
                if (borderType == BorderType.Cell && whichSide is "left" or "right")
                    if (space < 5.4m)
                        space = 5.4m;
                style.Add("padding-" + whichSide, space == 0 ? "0" : FormattableString.Invariant($"{space:0.0}pt"));
            }
            else
            {
                var sz = (int)side.Attribute(W.sz);
                var space = (decimal?)side.Attribute(W.space) ?? 0;
                var color = (string)side.Attribute(W.color);
                if (color is null or "auto")
                    color = "windowtext";
                else
                    color = ConvertColor(color);

                var borderWidthInPoints = Math.Max(1m, Math.Min(96m, Math.Max(2m, sz)) / 8m);

                var borderStyle = "solid";
                if (BorderStyleMap.TryGetValue(type, out var borderInfo))
                {
                    borderStyle = borderInfo.CssName;
                    if (type == "double")
                    {
                        borderWidthInPoints = sz switch
                        {
                            <= 8 => 2.5m,
                            <= 18 => 6.75m,
                            _ => sz / 3m,
                        };
                    }
                    else if (type == "triple")
                    {
                        borderWidthInPoints = sz switch
                        {
                            <= 8 => 8m,
                            <= 18 => 11.25m,
                            _ => 11.25m,
                        };
                    }
                    else if (type.Contains("dash", StringComparison.OrdinalIgnoreCase))
                    {
                        borderWidthInPoints = sz switch
                        {
                            <= 4 => 1m,
                            <= 12 => 1.5m,
                            _ => 2m,
                        };
                    }
                    else if (type != "single")
                        borderWidthInPoints = borderInfo.CssSize;
                }
                if (type is "outset" or "inset")
                    color = "";
                var borderWidth = FormattableString.Invariant($"{borderWidthInPoints:0.0}pt");

                style.Add("border-" + whichSide, borderStyle + " " + color + " " + borderWidth);
                if (borderType == BorderType.Cell && whichSide is "left" or "right")
                    if (space < 5.4m)
                        space = 5.4m;

                style.Add("padding-" + whichSide, space == 0 ? "0" : FormattableString.Invariant($"{space:0.0}pt"));
            }
        }

        private static readonly Dictionary<string, Func<string, string, string>> ShadeMapper = new()
        {
            { "auto", (c, f) => c },
            { "clear", (c, f) => f },
            { "nil", (c, f) => f },
            { "solid", (c, f) => c },
            { "diagCross", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "diagStripe", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "horzCross", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "horzStripe", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "pct10", (c, f) => ConvertColorFillPct(c, f, .1) },
            { "pct12", (c, f) => ConvertColorFillPct(c, f, .125) },
            { "pct15", (c, f) => ConvertColorFillPct(c, f, .15) },
            { "pct20", (c, f) => ConvertColorFillPct(c, f, .2) },
            { "pct25", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "pct30", (c, f) => ConvertColorFillPct(c, f, .3) },
            { "pct35", (c, f) => ConvertColorFillPct(c, f, .35) },
            { "pct37", (c, f) => ConvertColorFillPct(c, f, .375) },
            { "pct40", (c, f) => ConvertColorFillPct(c, f, .4) },
            { "pct45", (c, f) => ConvertColorFillPct(c, f, .45) },
            { "pct50", (c, f) => ConvertColorFillPct(c, f, .50) },
            { "pct55", (c, f) => ConvertColorFillPct(c, f, .55) },
            { "pct60", (c, f) => ConvertColorFillPct(c, f, .60) },
            { "pct62", (c, f) => ConvertColorFillPct(c, f, .625) },
            { "pct65", (c, f) => ConvertColorFillPct(c, f, .65) },
            { "pct70", (c, f) => ConvertColorFillPct(c, f, .7) },
            { "pct75", (c, f) => ConvertColorFillPct(c, f, .75) },
            { "pct80", (c, f) => ConvertColorFillPct(c, f, .8) },
            { "pct85", (c, f) => ConvertColorFillPct(c, f, .85) },
            { "pct87", (c, f) => ConvertColorFillPct(c, f, .875) },
            { "pct90", (c, f) => ConvertColorFillPct(c, f, .9) },
            { "pct95", (c, f) => ConvertColorFillPct(c, f, .95) },
            { "reverseDiagStripe", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "thinDiagCross", (c, f) => ConvertColorFillPct(c, f, .5) },
            { "thinDiagStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinHorzCross", (c, f) => ConvertColorFillPct(c, f, .3) },
            { "thinHorzStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinReverseDiagStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
            { "thinVertStripe", (c, f) => ConvertColorFillPct(c, f, .25) },
        };

        private static readonly Dictionary<string, string> ShadeCache = new();

        // fill is the background, color is the foreground
        private static string ConvertColorFillPct(string color, string fill, double pct)
        {
            if (color == "auto")
                color = "000000";
            if (fill == "auto")
                fill = "ffffff";
            var key = color + fill + pct.ToString(CultureInfo.InvariantCulture);
            if (ShadeCache.TryGetValue(key, out var cached))
                return cached;
            var fillRed = Convert.ToInt32(fill.Substring(0, 2), 16);
            var fillGreen = Convert.ToInt32(fill.Substring(2, 2), 16);
            var fillBlue = Convert.ToInt32(fill.Substring(4, 2), 16);
            var colorRed = Convert.ToInt32(color.Substring(0, 2), 16);
            var colorGreen = Convert.ToInt32(color.Substring(2, 2), 16);
            var colorBlue = Convert.ToInt32(color.Substring(4, 2), 16);
            var finalRed = (int)(fillRed - (fillRed - colorRed) * pct);
            var finalGreen = (int)(fillGreen - (fillGreen - colorGreen) * pct);
            var finalBlue = (int)(fillBlue - (fillBlue - colorBlue) * pct);
            var returnValue = $"{finalRed:x2}{finalGreen:x2}{finalBlue:x2}";
            ShadeCache.Add(key, returnValue);
            return returnValue;
        }

        private static void CreateStyleFromShd(Dictionary<string, string> style, XElement shd)
        {
            if (shd is null)
                return;
            var shadeType = (string)shd.Attribute(W.val);
            var color = (string)shd.Attribute(W.color);
            var fill = (string)shd.Attribute(W.fill);
            if (ShadeMapper.TryGetValue(shadeType, out var shadeFn))
            {
                color = shadeFn(color, fill);
            }
            if (color is not null)
            {
                var cvtColor = ConvertColor(color);
                if (!string.IsNullOrEmpty(cvtColor))
                    style.AddIfMissing("background", cvtColor);
            }
        }

        private static readonly Dictionary<string, string> NamedColors = new()
        {
            { "black", "black" },
            { "blue", "blue" },
            { "cyan", "aqua" },
            { "green", "green" },
            { "magenta", "fuchsia" },
            { "red", "red" },
            { "yellow", "yellow" },
            { "white", "white" },
            { "darkBlue", "#00008B" },
            { "darkCyan", "#008B8B" },
            { "darkGreen", "#006400" },
            { "darkMagenta", "#800080" },
            { "darkRed", "#8B0000" },
            { "darkYellow", "#808000" },
            { "darkGray", "#A9A9A9" },
            { "lightGray", "#D3D3D3" },
            { "none", "" },
        };

        private static void CreateColorProperty(string propertyName, string color, Dictionary<string, string> style)
        {
            if (color is null)
                return;

            // "auto" color is black for "color" and white for "background" property.
            if (color == "auto")
                color = propertyName == "color" ? "black" : "white";

            if (NamedColors.TryGetValue(color, out var namedColor1))
            {
                if (namedColor1 == "")
                    return;
                style.AddIfMissing(propertyName, namedColor1);
                return;
            }
            style.AddIfMissing(propertyName, "#" + color);
        }

        private static string ConvertColor(string color)
        {
            // "auto" color is black for "color" and white for "background" property.
            // As this method is only called for "background" colors, "auto" is translated
            // to "white" and never "black".
            if (color == "auto")
                color = "white";

            if (NamedColors.TryGetValue(color, out var namedColor2))
            {
                if (namedColor2 == "")
                    return "black";
                return namedColor2;
            }
            return "#" + color;
        }

        // Symbol/dingbat fonts that use non-standard character encodings. Directional marks
        // (LRM/RLM) must not be inserted into runs using these fonts as the marks render as
        // visible glyphs rather than invisible control characters.
        private static readonly FrozenSet<string> s_symbolFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Symbol",
            "Webdings",
            "Wingdings",
            "Wingdings2",
            "Wingdings 2",
            "Wingdings3",
            "Wingdings 3",
        }.ToFrozenSet(StringComparer.OrdinalIgnoreCase);

        private static readonly Dictionary<string, string> FontFallback = new()
        {
            // Sans-serif fonts
            { "Arial", @"'{0}', sans-serif" },
            { "Arial Narrow", @"'{0}', sans-serif" },
            { "Arial Rounded MT Bold", @"'{0}', sans-serif" },
            { "Arial Unicode MS", @"'{0}', sans-serif" },
            { "Berlin Sans FB", @"'{0}', sans-serif" },
            { "Berlin Sans FB Demi", @"'{0}', sans-serif" },
            { "Calibri", @"'{0}', sans-serif" },
            { "Calibri Light", @"'{0}', sans-serif" },
            { "Century Gothic", @"'{0}', sans-serif" },
            { "Comic Sans MS", @"'{0}', sans-serif" },
            { "Franklin Gothic Medium", @"'{0}', sans-serif" },
            { "Gill Sans MT", @"'{0}', sans-serif" },
            { "Gill Sans MT Condensed", @"'{0}', sans-serif" },
            { "Impact", @"'{0}', sans-serif" },
            { "Lucida Sans", @"'{0}', sans-serif" },
            { "Lucida Sans Unicode", @"'{0}', sans-serif" },
            { "Segoe UI", @"'{0}', sans-serif" },
            { "Segoe UI Light", @"'{0}', sans-serif" },
            { "Segoe UI Semibold", @"'{0}', sans-serif" },
            { "Tahoma", @"'{0}', sans-serif" },
            { "Trebuchet MS", @"'{0}', sans-serif" },
            { "Verdana", @"'{0}', sans-serif" },
            // Serif fonts
            { "Baskerville Old Face", @"'{0}', serif" },
            { "Book Antiqua", @"'{0}', serif" },
            { "Bookman Old Style", @"'{0}', serif" },
            { "Californian FB", @"'{0}', serif" },
            { "Cambria", @"'{0}', serif" },
            { "Constantia", @"'{0}', serif" },
            { "Garamond", @"'{0}', serif" },
            { "Georgia", @"'{0}', serif" },
            { "Lucida Bright", @"'{0}', serif" },
            { "Lucida Fax", @"'{0}', serif" },
            { "Palatino Linotype", @"'{0}', serif" },
            { "Times New Roman", @"'{0}', serif" },
            { "Wide Latin", @"'{0}', serif" },
            // Monospace fonts
            { "Consolas", @"'{0}', monospace" },
            { "Courier New", @"'{0}', monospace" },
            { "Lucida Console", @"'{0}', monospace" },
        };

        private static void CreateFontCssProperty(string font, Dictionary<string, string> style)
        {
            var normalizedFont = NormalizeFontFamilyWhitespace(font);
            if (FontFallback.TryGetValue(normalizedFont, out var fallbackFormat))
            {
                style.AddIfMissing("font-family", string.Format(fallbackFormat, normalizedFont));
                return;
            }

            // CSS font family names with whitespace can often be represented as a sequence of
            // identifiers without quotes, but names that cannot be represented safely as
            // identifiers (for example, those containing apostrophes) should be quoted. We
            // quote such names here for consistency and safety.
            var cssValue = NeedsCssQuoting(normalizedFont) ? QuoteCssString(normalizedFont) : normalizedFont;
            style.AddIfMissing("font-family", cssValue);
        }

        private static string NormalizeFontFamilyWhitespace(string font)
        {
            var trimmedFont = font.Trim();
            if (trimmedFont.Length == 0)
            {
                return trimmedFont;
            }

            var sb = new StringBuilder(trimmedFont.Length);
            var inWhitespace = false;
            foreach (var c in trimmedFont)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (!inWhitespace)
                    {
                        sb.Append(' ');
                        inWhitespace = true;
                    }
                }
                else
                {
                    sb.Append(c);
                    inWhitespace = false;
                }
            }

            return sb.ToString();
        }

        // Returns true when the font family name must be quoted in CSS.
        // A name can be unquoted only when it is a valid CSS identifier:
        // the first character must be a valid identifier start character and the
        // remaining characters must be valid identifier continuation characters.
        // ASCII punctuation such as apostrophes, and names that start with digits,
        // must be written as quoted CSS strings.
        private static bool IsValidCssIdentifierStart(char c) => c >= 0x80 || char.IsLetter(c) || c == '-' || c == '_';

        private static bool IsValidCssIdentifierPart(char c) =>
            c >= 0x80 || char.IsLetterOrDigit(c) || c == '-' || c == '_';

        private static bool NeedsCssQuoting(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return true;
            }

            if (!IsValidCssIdentifierStart(value[0]))
            {
                return true;
            }

            for (var i = 1; i < value.Length; i++)
            {
                var c = value[i];
                if (char.IsWhiteSpace(c) || !IsValidCssIdentifierPart(c))
                {
                    return true;
                }
            }

            return false;
        }

        private static string QuoteCssString(string value)
        {
            var quote = value.Contains('\'') && !value.Contains('"') ? '"' : '\'';
            return $"{quote}{EscapeCssString(value, quote)}{quote}";
        }

        private static string EscapeCssString(string value, char quote)
        {
            var sb = new StringBuilder(value.Length);
            foreach (var c in value)
            {
                switch (c)
                {
                    case '\\':
                        sb.Append(@"\\");
                        break;
                    case '\n' or '\r' or '\f':
                        sb.Append('\\');
                        sb.Append(((int)c).ToString("x", CultureInfo.InvariantCulture));
                        sb.Append(' ');
                        break;
                    default:
                        if (c == quote)
                        {
                            sb.Append('\\');
                        }

                        sb.Append(c);
                        break;
                }
            }

            return sb.ToString();
        }

        private static bool GetBoolProp(XElement runProps, XName xName)
        {
            var p = runProps.Element(xName);
            if (p is null)
                return false;
            var v = p.Attribute(W.val);
            if (v is null)
                return true;
            var s = v.Value;
            return s switch
            {
                "0" => false,
                "1" => true,
                _ when s.Equals("false", StringComparison.OrdinalIgnoreCase) => false,
                _ when s.Equals("true", StringComparison.OrdinalIgnoreCase) => true,
                _ => false,
            };
        }

        private static object ConvertContentThatCanContainFields(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            IEnumerable<XElement> elements
        )
        {
            var grouped = elements
                .GroupAdjacent(e =>
                {
                    var stack = e.Annotation<Stack<FieldRetriever.FieldElementTypeInfo>>();
                    return stack is null || !stack.Any() ? (int?)null : stack.Select(st => st.Id).Min();
                })
                .ToList();

            var txformed = grouped
                .Select(g =>
                {
                    var key = g.Key;
                    if (key is null)
                        return (object)g.Select(n => ConvertToHtmlTransform(wordDoc, settings, n, false, 0m));

                    var instrText = FieldRetriever
                        .InstrText(g.First().Ancestors().Last(), (int)key)
                        .TrimStart('{')
                        .TrimEnd('}');

                    var parsed = FieldRetriever.ParseField(instrText);
                    if (parsed.FieldType != "HYPERLINK")
                        return g.Select(n => ConvertToHtmlTransform(wordDoc, settings, n, false, 0m));

                    var content = g.DescendantsAndSelf(W.r).Select(run => ConvertRun(wordDoc, settings, run));
                    var a =
                        parsed.Arguments.Length > 0
                            ? new XElement(Xhtml.a, new XAttribute("href", parsed.Arguments[0]), content)
                            : new XElement(Xhtml.a, content);
                    if (!a.Nodes().Any())
                        a.Add(new XText(""));
                    return a;
                })
                .ToList();

            return txformed;
        }

        #region Text Box Processing

        private static XElement? ProcessTextBoxDrawing(
            WordprocessingDocument wordDoc,
            WmlToHtmlConverterSettings settings,
            XElement drawingElement
        )
        {
            var containerElement = drawingElement
                .Elements()
                .FirstOrDefault(e => e.Name == WP.inline || e.Name == WP.anchor);
            if (containerElement is null)
                return null;

            var txbx = containerElement
                .Elements(A.graphic)
                .Elements(A.graphicData)
                .Elements(WPS.wsp)
                .Elements(WPS.txbx)
                .FirstOrDefault();
            if (txbx is null)
                return null;

            var txbxContent = txbx.Element(W.txbxContent);
            if (txbxContent is null)
                return null;

            var extentCx = (int?)containerElement.Elements(WP.extent).Attributes(NoNamespace.cx).FirstOrDefault();
            var extentCy = (int?)containerElement.Elements(WP.extent).Attributes(NoNamespace.cy).FirstOrDefault();

            var style = new Dictionary<string, string>();
            style.AddIfMissing("display", "inline-block");
            style.AddIfMissing("overflow", "hidden");
            style.AddIfMissing("padding", "2pt");
            if (extentCx is not null)
                style.AddIfMissing(
                    "width",
                    FormattableString.Invariant($"{(float)extentCx / ImageInfo.EmusPerInch:0.00}in")
                );
            if (extentCy is not null)
                style.AddIfMissing(
                    "min-height",
                    FormattableString.Invariant($"{(float)extentCy / ImageInfo.EmusPerInch:0.00}in")
                );

            // Only float anchored text boxes when the wrap mode implies surrounding text should flow
            // around the shape. wp:wrapNone means no text wrap (overlap), and wp:wrapTopAndBottom
            // pushes the shape to its own line — neither needs float.
            if (containerElement.Name == WP.anchor)
            {
                var hasTextWrapping = containerElement
                    .Elements()
                    .Any(e => e.Name == WP.wrapSquare || e.Name == WP.wrapTight || e.Name == WP.wrapThrough);
                if (hasTextWrapping)
                    style.AddIfMissing("float", "left");
            }

            // Text boxes have independent layout — reset the outer paragraph margin so that list/indent
            // offsets from the surrounding context do not incorrectly bleed into the text box interior.
            var content = txbxContent
                .Elements()
                .Select(e => ConvertToHtmlTransform(wordDoc, settings, e, false, 0m))
                .ToList();

            var div = new XElement(Xhtml.div, content);
            div.AddAnnotation(style);
            return div;
        }

        #endregion

        #region Image Processing

        // Don't process wmf files (with contentType == "image/x-wmf") because GDI consumes huge amounts
        // of memory when dealing with wmf perhaps because it loads a DLL to do the rendering?
        // It actually works, but is not recommended.
        private static readonly List<string> ImageContentTypes = new()
        {
            "image/png",
            "image/gif",
            "image/tiff",
            "image/jpeg",
        };

        public static XElement ProcessImage(
            WordprocessingDocument wordDoc,
            XElement element,
            Func<ImageInfo, XElement> imageHandler
        )
        {
            if (imageHandler is null)
            {
                return null;
            }
            if (element.Name == W.drawing)
            {
                return ProcessDrawing(wordDoc, element, imageHandler);
            }
            if (element.Name == W.pict || element.Name == W._object)
            {
                return ProcessPictureOrObject(wordDoc, element, imageHandler);
            }
            return null;
        }

        private static XElement ProcessDrawing(
            WordprocessingDocument wordDoc,
            XElement element,
            Func<ImageInfo, XElement> imageHandler
        )
        {
            var containerElement = element.Elements().FirstOrDefault(e => e.Name == WP.inline || e.Name == WP.anchor);
            if (containerElement is null)
                return null;

            string hyperlinkUri = null;
            var hyperlinkElement = element
                .Elements(WP.inline)
                .Elements(WP.docPr)
                .Elements(A.hlinkClick)
                .FirstOrDefault();
            if (hyperlinkElement is not null)
            {
                var rId = (string)hyperlinkElement.Attribute(R.id);
                if (rId is not null)
                {
                    var hyperlinkRel = wordDoc.MainDocumentPart.HyperlinkRelationships.FirstOrDefault(hlr =>
                        hlr.Id == rId
                    );
                    if (hyperlinkRel is not null)
                    {
                        hyperlinkUri = hyperlinkRel.Uri.ToString();
                    }
                }
            }

            var extentCx = (int?)containerElement.Elements(WP.extent).Attributes(NoNamespace.cx).FirstOrDefault();
            var extentCy = (int?)containerElement.Elements(WP.extent).Attributes(NoNamespace.cy).FirstOrDefault();
            var altText =
                (string)containerElement.Elements(WP.docPr).Attributes(NoNamespace.descr).FirstOrDefault()
                ?? ((string)containerElement.Elements(WP.docPr).Attributes(NoNamespace.name).FirstOrDefault() ?? "");

            var blipFill = containerElement
                .Elements(A.graphic)
                .Elements(A.graphicData)
                .Elements(Pic._pic)
                .Elements(Pic.blipFill)
                .FirstOrDefault();
            if (blipFill is null)
                return null;

            var imageRid = (string)blipFill.Elements(A.blip).Attributes(R.embed).FirstOrDefault();
            if (imageRid is null)
                return null;

            var pp3 = wordDoc.MainDocumentPart.Parts.FirstOrDefault(pp => pp.RelationshipId == imageRid);
            if (pp3 == default)
                return null;

            var imagePart = (ImagePart)pp3.OpenXmlPart;
            if (imagePart is null)
                return null;

            // If the image markup points to a NULL image, then following will throw an ArgumentOutOfRangeException
            try
            {
                imagePart = (ImagePart)wordDoc.MainDocumentPart.GetPartById(imageRid);
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }

            var contentType = imagePart.ContentType;
            if (!ImageContentTypes.Contains(contentType))
                return null;

            using var partStream = imagePart.GetStream();
            using var image = SKBitmap.Decode(partStream);
            if (image is null)
                return null;
            if (extentCx is not null && extentCy is not null)
            {
                var imageInfo = new ImageInfo()
                {
                    Image = image,
                    ImgStyleAttribute = new XAttribute(
                        "style",
                        FormattableString.Invariant(
                            $"width: {(float)extentCx / ImageInfo.EmusPerInch}in; height: {(float)extentCy / ImageInfo.EmusPerInch}in"
                        )
                    ),
                    ContentType = contentType,
                    DrawingElement = element,
                    AltText = altText,
                };
                var imgElement2 = imageHandler(imageInfo);
                if (hyperlinkUri is not null)
                {
                    return new XElement(
                        XhtmlNoNamespace.a,
                        new XAttribute(XhtmlNoNamespace.href, hyperlinkUri),
                        imgElement2
                    );
                }
                return imgElement2;
            }

            var imageInfo2 = new ImageInfo()
            {
                Image = image,
                ContentType = contentType,
                DrawingElement = element,
                AltText = altText,
            };
            var imgElement = imageHandler(imageInfo2);
            if (hyperlinkUri is not null)
            {
                return new XElement(
                    XhtmlNoNamespace.a,
                    new XAttribute(XhtmlNoNamespace.href, hyperlinkUri),
                    imgElement
                );
            }
            return imgElement;
        }

        private static XElement ProcessPictureOrObject(
            WordprocessingDocument wordDoc,
            XElement element,
            Func<ImageInfo, XElement> imageHandler
        )
        {
            var imageRid = (string)
                element.Elements(VML.shape).Elements(VML.imagedata).Attributes(R.id).FirstOrDefault();
            if (imageRid is null)
                return null;

            try
            {
                var pp = wordDoc.MainDocumentPart.Parts.FirstOrDefault(pp2 => pp2.RelationshipId == imageRid);
                if (pp == default)
                    return null;

                var imagePart = (ImagePart)pp.OpenXmlPart;
                if (imagePart == default)
                    return null;

                var contentType = imagePart.ContentType;
                if (!ImageContentTypes.Contains(contentType))
                    return null;

                using var partStream = imagePart.GetStream();
                try
                {
                    using var bitmap = SKBitmap.Decode(partStream);
                    if (bitmap is null)
                        return null;

                    var imageInfo = new ImageInfo
                    {
                        Image = bitmap,
                        ContentType = contentType,
                        DrawingElement = element,
                    };

                    var style = (string)element.Elements(VML.shape).Attributes("style").FirstOrDefault();
                    if (style is null)
                        return imageHandler(imageInfo);

                    var tokens = style.Split(';');
                    var widthInPoints = WidthInPoints(tokens);
                    var heightInPoints = HeightInPoints(tokens);
                    if (widthInPoints is not null && heightInPoints is not null)
                    {
                        imageInfo.ImgStyleAttribute = new XAttribute(
                            "style",
                            FormattableString.Invariant($"width: {widthInPoints}pt; height: {heightInPoints}pt")
                        );
                    }
                    return imageHandler(imageInfo);
                }
                catch (OutOfMemoryException)
                {
                    // the Bitmap class can throw OutOfMemoryException, which means the bitmap is messed up, so punt.
                    return null;
                }
                catch (ArgumentException)
                {
                    return null;
                }
            }
            catch (ArgumentOutOfRangeException)
            {
                return null;
            }
        }

        private static float? HeightInPoints(IEnumerable<string> tokens)
        {
            return SizeInPoints(tokens, "height");
        }

        private static float? WidthInPoints(IEnumerable<string> tokens)
        {
            return SizeInPoints(tokens, "width");
        }

        private static float? SizeInPoints(IEnumerable<string> tokens, string name)
        {
            var sizeString = tokens
                .Select(t => new { Name = t.Split(':').First(), Value = t.Split(':').Skip(1).Take(1).FirstOrDefault() })
                .Where(p => p.Name == name)
                .Select(p => p.Value)
                .FirstOrDefault();

            if (sizeString is not null && sizeString.Length > 2 && sizeString.Substring(sizeString.Length - 2) == "pt")
            {
                if (float.TryParse(sizeString.Substring(0, sizeString.Length - 2), out var size))
                    return size;
            }
            return null;
        }

        #endregion
    }

    public static class HtmlConverterExtensions
    {
        public static void AddIfMissing(this Dictionary<string, string> style, string propName, string value)
        {
            style.TryAdd(propName, value);
        }
    }
}
