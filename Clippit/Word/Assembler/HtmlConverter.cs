using System.Collections;
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
        private static readonly HtmlToWmlConverterSettings htmlConverterSettings =
            HtmlToWmlConverter.GetDefaultSettings();

        private static readonly Regex detectEntityRegEx = new Regex("^&(?:#([0-9]+)|#x([0-9a-fA-F]+)|([0-9a-zA-Z]+));");

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
        internal static IEnumerable<object> ProcessContentElement(
            this XElement element,
            XElement data,
            TemplateError templateError,
            ref OpenXmlPart part
        )
        {
            var xPath = (string)element.Attribute(PA.Select);
            var optionalString = (string)element.Attribute(PA.Optional);
            bool optional = (optionalString != null && optionalString.ToLower() == "true");

            string[] values = data.EvaluateXPath(xPath, optional);

            // if we no data returned then just return an empty run
            if (values.Length == 0)
            {
                return new[] { new XElement(W.r, W.t) };
            }

            // otherwise split the values if there are new line characters
            values = values
                .SelectMany(x => x.Replace("\r\n", "\n", StringComparison.OrdinalIgnoreCase).Split('\n'))
                .ToArray();

            List<object> results = new List<object>();
            for (int i = 0; i < values.Length; i++)
            {
                // try processing as XML
                XElement parsedElement = XElement.Parse($"<xhtml>{EscapeAmpersands(values[i])}</xhtml>");

                results.Add(
                    Transform(
                        parsedElement,
                        htmlConverterSettings,
                        part,
                        i == 0 ? NextExpected.Run : NextExpected.Paragraph,
                        true
                    )
                );
            }

            results = FlattenResults(results);

            if (results.Count == 0)
            {
                return new[] { new XElement(W.r, W.t) };
            }

            return results;
        }

        private static List<object> FlattenResults(IEnumerable content)
        {
            // flatten the returned content
            List<object> results = new List<object>();
            foreach (object obj in content)
            {
                if (obj is IEnumerable)
                {
                    results.AddRange(FlattenResults(obj as IEnumerable));
                }
                else
                {
                    results.Add(obj);
                }
            }

            return results;
        }

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
                    var rId = Relationships.GetNewRelationshipId();
                    var href = (string)element.Attribute(NoNamespace.href);
                    if (href != null)
                    {
                        Uri uri = null;
                        try
                        {
                            uri = href.GetUri();
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

                            if (nextExpected == NextExpected.Paragraph)
                            {
                                return new XElement(W.p, hyperlink);
                            }

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
                                .Any(d => d.Name == XhtmlNoNamespace.li || d.Name == XhtmlNoNamespace.p)
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
                        return new XElement(W.r);
                    }

                // if no match up to this point, then just recursively process descendants
                return element.Nodes().Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace));
            }

            // process text nodes unless their parent is a title tag
            if (node.Parent.Name != XhtmlNoNamespace.title)
                return GenerateNextExpected(node, settings, part, null, nextExpected, preserveWhiteSpace);

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
                            new XElement(
                                W.r,
                                HtmlToWmlConverterCore.GetRunProperties(xTextNode, settings),
                                new XElement(
                                    W.t,
                                    HtmlToWmlConverterCore.GetXmlSpaceAttribute(textNodeString),
                                    textNodeString
                                )
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
                    return element
                        .Nodes()
                        .Select(n => Transform(n, settings, part, nextExpected, preserveWhiteSpace))
                        .AsEnumerable();
                }
                else
                {
                    var textNodeString = HtmlToWmlConverterCore.GetDisplayText((XText)node, preserveWhiteSpace);
                    var rPr = HtmlToWmlConverterCore.GetRunProperties((XText)node, settings);
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
