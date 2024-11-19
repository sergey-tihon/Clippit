using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

using Clippit.Html;
using Clippit.Internal;
using DocumentFormat.OpenXml.Packaging;

using NextExpected = Clippit.Html.HtmlToWmlConverterCore.NextExpected;

namespace Clippit.Word.Assembler
{
    internal static class HtmlConverter
    {
        private static readonly HtmlToWmlConverterSettings htmlConverterSettings = HtmlToWmlConverter.GetDefaultSettings();

        private static readonly Regex detectEntityRegEx = new Regex("^&(?:#([0-9]+)|#x([0-9a-fA-F]+)|([0-9a-zA-Z]+));");

        private static readonly string[] strongElements = { "b", "strong" };
        private static readonly string[] emphasisElements = { "i", "em" };
        private static readonly string[] underlineElements = { "u" };
        private static readonly string[] hyperlinkElements = { "a" };

        private static readonly string[] newlineInlineElements = { "br" };
        private static readonly string[] newlineBlockElements = { "p", "div" };

        /// <summary>
        /// Method processes a string that contains inline html tags and generates a run with the necessary properties
        /// Supported inline html tags: b, i, em, strong, u, br, a
        /// Supported block tags: p, div
        /// TODO: add support for the following html tags: big, small, sub, sup, span.
        /// </summary>
        /// <param name="element">Source element.</param>
        /// <param name="data">Data element with content.</param>
        /// <param name="pPr">The paragraph properties.</param>
        /// <param name="templateError">Error indicator.</param>
        internal static object ProcessContentElement(
            this XElement element,
            XElement data,
            XElement pPr,
            TemplateError templateError,
            ref OpenXmlPart part
        )
        {
            var xPath = (string)element.Attribute(PA.Select);
            var optionalString = (string)element.Attribute(PA.Optional);
            bool optional = (optionalString != null && optionalString.ToLower() == "true");

            return data.EvaluateXPath(xPath, optional)
                .Select(x =>
                    Transform(XElement.Parse(x),
                        htmlConverterSettings,
                        null,
                        NextExpected.Paragraph,
                        true));
            
            /*bool bold = false,
                italic = false,
                underline = false,
                link = false;

            // create an empty paragraph with the same properties as the current paragraph in case we have new lines
            XElement emptyPara = new XElement(W.p, pPr);

            // initialise the list of paragraphs to return
            List<XElement> elementList = new List<XElement>();
            try
            {
                // create the xml reader
                using (StringReader stringReader = new StringReader(value))
                {
                    using (XmlReader xmlReader = XmlReader.Create(stringReader))
                    {
                        string content = string.Empty;
                        string url = string.Empty;
                        while (xmlReader.Read())
                        {
                            string tagName = xmlReader.LocalName.ToLower();

                            // process XML elements
                            if (
                                xmlReader.NodeType == XmlNodeType.Element
                                || xmlReader.NodeType == XmlNodeType.EndElement
                            )
                            {
                                bool isStart = xmlReader.IsStartElement();
                                bool isEmpty = xmlReader.IsEmptyElement;

                                if (hyperlinkElements.Contains(tagName) && link != isStart)
                                {
                                    link = isStart;
                                }

                                if (link && hyperlinkElements.Contains(tagName))
                                {
                                    url = xmlReader.GetAttribute("href");
                                }

                                if (strongElements.Contains(tagName) && bold != isStart)
                                {
                                    bold = isStart;
                                }
                                else if (emphasisElements.Contains(tagName) && italic != isStart)
                                {
                                    italic = isStart;
                                }
                                else if (underlineElements.Contains(tagName) && underline != isStart)
                                {
                                    underline = isStart;
                                }
                                else if (
                                    (
                                        newlineInlineElements.Contains(tagName)
                                        || (newlineBlockElements.Contains(tagName)) && (!isStart || isEmpty)
                                    )
                                )
                                {
                                    AppendRun(elementList, CreateRun(content, bold, italic, underline, ref part, url));
                                    content = string.Empty;
                                    url = string.Empty;

                                    // create a new paragraph for the remaining content
                                    elementList.Add(new XElement(emptyPara));
                                }

                                continue; // ignore other elements
                            }

                            // process text and whitespace
                            if (xmlReader.NodeType == XmlNodeType.Text || xmlReader.NodeType == XmlNodeType.Whitespace)
                            {
                                content += xmlReader.Value;

                                AppendRun(elementList, CreateRun(content, bold, italic, underline, ref part, url));
                                content = string.Empty;
                                url = string.Empty;
                            }
                        }

                        if (!string.IsNullOrEmpty(content)) // append the last run
                        {
                            AppendRun(elementList, CreateRun(content, bold, italic, underline, ref part));
                        }
                    }
                }
            }
            catch (Exception e)
            {
                return element.CreateContextErrorMessage("XPathException: " + e.Message, templateError);
            }

            // if we have trailing empty paragraphs lets remove them from the list
            while (
                elementList.Count > 0
                && elementList.Last().Name == W.p
                && elementList.Last().Descendants().Where(x => x.Name == W.t).Any() == false
            )
            {
                elementList.RemoveAt(elementList.Count - 1);
            }

            return elementList;*/
        }

        private static void AppendRun(List<XElement> elementList, XElement newRun)
        {
            if (elementList.Count == 0 || elementList.Last().Name != W.p)
            {
                // if we have not created a new paragraph yet then add the run to the list
                elementList.Add(newRun);
            }
            else
            {
                // otherwise add our run to the last paragraph
                elementList.Last().Add(newRun);
            }
        }

        private static XElement CreateRun(
            string text,
            bool bold,
            bool italic,
            bool underline,
            ref OpenXmlPart part,
            string url = ""
        )
        {
            if (!string.IsNullOrWhiteSpace(url))
            {
                XElement hyp = new XElement(W.hyperlink);
                XNamespace idNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                HyperlinkRelationship rel = null;

                var mainDocumentPart = part as MainDocumentPart;
                if (mainDocumentPart != null)
                {
                    rel = mainDocumentPart.AddHyperlinkRelationship(url.GetUri(), true);
                }

                XElement relId = new XElement(idNs + "id", rel.Id);
                hyp.SetAttributeValue(relId.Name, relId.Value);

                XElement r2 = new XElement(W.r);

                // add the required properties according to the formatting which is currently on
                XElement rProps = new XElement(W.rPr, new XElement(W.rStyle, new XAttribute(W.val, "Hyperlink")));

                if (bold)
                {
                    rProps.Add(new XElement(W.b));
                    rProps.Add(new XElement(W.bCs));
                }
                if (italic)
                {
                    rProps.Add(new XElement(W.i));
                    rProps.Add(new XElement(W.iCs));
                }
                if (underline)
                {
                    rProps.Add(new XElement(W.u, new XAttribute(W.val, "single")));
                }

                r2.Add(rProps);

                if (!string.IsNullOrEmpty(text) && text.Length > 0)
                {
                    r2.Add(
                        new XElement(
                            W.t,
                            PreserveWhitespace(text) ? new XAttribute(XNamespace.Xml + "space", "preserve") : null,
                            text
                        )
                    );
                }

                hyp.Add(r2);

                return hyp;
            }

            // create a run
            XElement r = new XElement(W.r);

            // add the required properties according to the formatting which is currently on
            if (bold || italic || underline)
            {
                r.Add(
                    new XElement(
                        W.rPr,
                        bold ? new XElement(W.b) : null,
                        italic ? new XElement(W.i) : null,
                        underline ? new XElement(W.u, new XAttribute(W.val, "single")) : null
                    )
                );
            }

            if (!string.IsNullOrEmpty(text) && text.Length > 0)
            {
                r.Add(
                    new XElement(
                        W.t,
                        PreserveWhitespace(text) ? new XAttribute(XNamespace.Xml + "space", "preserve") : null,
                        text ?? string.Empty
                    )
                );
            }

            return r;
        }

        private static bool PreserveWhitespace(string text)
        {
            if (text != null && text.Length > 0)
            {
                // do we need to preserve whitespace?
                char firstChar = text[0];
                char lastChar = text[text.Length - 1];

                return char.IsWhiteSpace(firstChar) || char.IsWhiteSpace(lastChar) || string.IsNullOrWhiteSpace(text);
            }

            return false;
        }

        /* SCRATCH */
        private static object Transform(
            XNode node,
            HtmlToWmlConverterSettings settings,
            OpenXmlPart part,
            NextExpected nextExpected,
            bool preserveWhiteSpace
        )
        {
            var element = node as XElement;
            if (element != null)
            {
                if (element.Name == XhtmlNoNamespace.a)
                {
                    var rId = DocumentAssembler.GetNextRelationshipId(part);
                    var href = (string)element.Attribute(NoNamespace.href);
                    if (href != null)
                    {
                        Uri uri = null;
                        try
                        {
                            uri = new Uri(href);
                        }
                        catch (UriFormatException)
                        {
                            var rPr = HtmlToWmlConverterCore.GetRunProperties(element, settings);
                            var run = new XElement(W.r, rPr, new XElement(W.t, element.Value));
                            return new[] { run };
                        }

                        if (uri != null)
                        {
                            part.AddHyperlinkRelationship(uri, true, rId);
                            if (element.Element(XhtmlNoNamespace.img) != null)
                            {
                                var imageTransformed = element
                                    .Nodes()
                                    .Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace))
                                    .OfType<XElement>();
                                var newImageTransformed = imageTransformed
                                    .Select(i =>
                                    {
                                        if (i.Elements(W.drawing).Any())
                                        {
                                            var newRun = new XElement(i);
                                            var docPr = newRun
                                                .Elements(W.drawing)
                                                .Elements(WP.inline)
                                                .Elements(WP.docPr)
                                                .FirstOrDefault();
                                            if (docPr != null)
                                            {
                                                var hlinkClick = new XElement(
                                                    A.hlinkClick,
                                                    new XAttribute(R.id, rId),
                                                    new XAttribute(XNamespace.Xmlns + "a", A.a.NamespaceName)
                                                );
                                                docPr.Add(hlinkClick);
                                            }
                                            return newRun;
                                        }
                                        return i;
                                    })
                                    .ToList();
                                return newImageTransformed;
                            }

                            var rPr = HtmlToWmlConverterCore.GetRunProperties(element, settings);
                            var hyperlink = new XElement(
                                W.hyperlink,
                                new XAttribute(R.id, rId),
                                new XElement(W.r, rPr, new XElement(W.t, element.Value))
                            );
                            return new[] { hyperlink };
                        }
                    }
                    return null;
                }

                if (element.Name == XhtmlNoNamespace.b)
                    return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.div)
                {
                    if (nextExpected == NextExpected.Paragraph)
                    {
                        if (
                            element
                                .Descendants()
                                .Any(d =>
                                    d.Name == XhtmlNoNamespace.li ||
                                    d.Name == XhtmlNoNamespace.p
                                )
                        )
                        {
                            return element
                                .Nodes()
                                .Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));
                        }
                        else
                        {
                            return GenerateNextExpected(element, settings, part, null, nextExpected, false);
                        }
                    }
                    else
                    {
                        return element
                            .Nodes()
                            .Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));
                    }
                }

                if (element.Name == XhtmlNoNamespace.em)
                    return element
                        .Nodes()
                        .Select(n => Transform(n, settings, part, NextExpected.Run, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.html)
                    return element
                        .Nodes()
                        .Select(n => Transform(n, settings, part, NextExpected.Paragraph, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.i)
                    return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.li)
                {
                    return GenerateNextExpected(element, settings, part, null, NextExpected.Paragraph, false);
                }

                if (element.Name == XhtmlNoNamespace.ol)
                    return element
                        .Nodes()
                        .Select(n => Transform(n, settings, part, NextExpected.Paragraph, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.p)
                {
                    return GenerateNextExpected(element, settings, part, null, NextExpected.Paragraph, false);
                }

                if (element.Name == XhtmlNoNamespace.strong)
                    return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.sub)
                    return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.sup)
                    return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.u)
                    return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.ul)
                    return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));

                if (element.Name == XhtmlNoNamespace.br)
                    if (nextExpected == NextExpected.Paragraph)
                    {
                        return new XElement(W.p, new XElement(W.r, new XElement(W.t)));
                    }
                    else
                    {
                        return new XElement(W.r, new XElement(W.br));
                    }

                // if no match up to this point, then just recursively process descendants
                return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));
            }

            return null;
        }

        private static object GenerateNextExpected(
            XNode node,
            HtmlToWmlConverterSettings settings,
            OpenXmlPart part,
            string styleName,
            NextExpected nextExpected,
            bool preserveWhiteSpace
        )
        {
            if (nextExpected == NextExpected.Paragraph)
            {
                var element = node as XElement;
                if (element != null)
                {
                    return new XElement(
                        W.p,
                        HtmlToWmlConverterCore.GetParagraphProperties(element, styleName, settings),
                        element.Nodes().Select(n => Transform(n, settings, part, NextExpected.Run, preserveWhiteSpace))
                    );
                }
                else
                {
                    var xTextNode = node as XText;
                    if (xTextNode != null)
                    {
                        var textNodeString = HtmlToWmlConverterCore.GetDisplayText(xTextNode, preserveWhiteSpace);
                        var p = new XElement(
                            W.p,
                            HtmlToWmlConverterCore.GetParagraphProperties(node.Parent, null, settings),
                            new XElement(
                                W.r,
                                HtmlToWmlConverterCore.GetRunProperties((XElement)node, settings),
                                new XElement(W.t, HtmlToWmlConverterCore.GetXmlSpaceAttribute(textNodeString), textNodeString)
                            )
                        );
                        return p;
                    }
                    return null;
                }
            }
            else
            {
                var element = node as XElement;
                if (element != null)
                {
                    return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));
                }
                else
                {
                    var textNodeString = HtmlToWmlConverterCore.GetDisplayText((XText)node, preserveWhiteSpace);
                    var rPr = HtmlToWmlConverterCore.GetRunProperties((XElement)node, settings);
                    var r = new XElement(
                        W.r,
                        rPr,
                        new XElement(W.t, HtmlToWmlConverterCore.GetXmlSpaceAttribute(textNodeString), textNodeString)
                    );
                    return r;
                }
            }
        }


        private static string EscapeAmpersands(string value)
        {
            // check whether we have any processing to do
            if (!string.IsNullOrWhiteSpace(value) && value.Contains('&', StringComparison.OrdinalIgnoreCase))
            {
                string result = string.Empty;

                int ampIndex = value.IndexOf('&', StringComparison.OrdinalIgnoreCase);
                while (ampIndex >= 0)
                {
                    // put everything before the ampersand into the result
                    result += value.Substring(0, ampIndex);

                    // then trim the value back
                    value = value.Substring(ampIndex);

                    // now check whether ampersand we have found is the start of an entity
                    Match m = detectEntityRegEx.Match(value);
                    if (m.Success)
                    {
                        // if this is an entity then add to result
                        result += value.Substring(0, m.Length);

                        // then remove entity from input
                        value = value.Substring(m.Length);
                    }
                    else
                    {
                        // add escaped ampersand to result
                        result += "&amp;";

                        // then remove ampersand from input
                        value = value.Substring(1);
                    }

                    ampIndex = value.IndexOf('&', StringComparison.OrdinalIgnoreCase);
                }

                // add any remaining string
                if (!string.IsNullOrEmpty(value))
                {
                    result += value;
                }

                return result;
            }

            return value;
        }
    }
}
