using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;
using HtmlAgilityPack;

namespace Clippit.Word.Assembler
{
    internal static class HtmlConverter
    {
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

            string value;
            try
            {
                string[] values = data.EvaluateXPath(xPath, optional);

                value = PreprocessHtml(EscapeAmpersands(string.Join("\r\n", values)));
            }
            catch (XPathException e)
            {
                return element.CreateContextErrorMessage("XPathException: " + e.Message, templateError);
            }

            bool bold = false,
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
                && elementList.Last().Descendants().Where(x => x.Name == W.t).Count() == 0
            )
            {
                elementList.RemoveAt(elementList.Count - 1);
            }

            return elementList;
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

                if (!string.IsNullOrEmpty(text) && text.Any())
                {
                    r2.Add(
                        new XElement(
                            W.t,
                            PreserveWhitespace(text) ? new XAttribute(XNamespace.Xml + "space", "preserve") : null,
                            text ?? string.Empty
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

            if (!string.IsNullOrEmpty(text) && text.Any())
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
            if (text != null && text.Any())
            {
                // do we need to preserve whitespace?
                char firstChar = text[0];
                char lastChar = text[text.Length - 1];

                return char.IsWhiteSpace(firstChar) || char.IsWhiteSpace(lastChar) || string.IsNullOrWhiteSpace(text);
            }

            return false;
        }

        private static string PreprocessHtml(string value)
        {
            value = HtmlNormalizeLineBreaks(value);
            value = HtmlExpandEntities(value);
            return HtmlToXml(value);
        }

        private static string HtmlNormalizeLineBreaks(string value)
        {
            // assume our text is either text with inline html tags or plain text
            // wrap in html tags and add br tags in place of crlf or lf
            return value.Replace("\r\n", "<br/>").Replace("\n", "<br/>").Replace("<br>", "<br/>");
        }

        private static string HtmlExpandEntities(string value)
        {
            return WebUtility.HtmlDecode(value);
        }

        private static string HtmlToXml(string value)
        {
            try
            {
                HtmlDocument doc = new HtmlDocument();
                doc.OptionOutputAsXml = true;
                doc.OptionWriteEmptyNodes = true; //autoclose hr, br etc
                doc.LoadHtml($"<html>{value}</html>");

                // remove breaks that are in paragraphs that have no content
                // as this will end up creating two paragraphs in Word one for the <p> and one for the <br>
                var brs = doc.DocumentNode.SelectNodes("//br[ancestor::p[not(.//text()[normalize-space()])]]");
                if (brs != null)
                {
                    foreach (var br in brs)
                    {
                        br.Remove();
                    }
                }

                StringBuilder sb = new StringBuilder();
                XmlWriterSettings xmlWtrSettings = new XmlWriterSettings() { OmitXmlDeclaration = true };
                using (XmlWriter xmlWtr = XmlWriter.Create(sb, xmlWtrSettings))
                {
                    doc.Save(xmlWtr);
                    xmlWtr.Flush();
                }

                return sb.ToString();
            }
            catch
            {
                throw;
            }
        }

        private static string EscapeAmpersands(string value)
        {
            // check whether we have any processing to do
            if (!string.IsNullOrWhiteSpace(value) && value.Contains("&"))
            {
                string result = string.Empty;

                int ampIndex = value.IndexOf("&");
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

                    ampIndex = value.IndexOf('&');
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
