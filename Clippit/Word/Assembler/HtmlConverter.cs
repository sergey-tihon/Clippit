using System.Collections;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Clippit.Html;
using Clippit.Internal;
using DocumentFormat.OpenXml.Packaging;
using SixLabors.ImageSharp;
using NextExpected = Clippit.Html.HtmlToWmlConverterCore.NextExpected;

namespace Clippit.Word.Assembler
{
    internal static class HtmlConverter
    {
        private static readonly HtmlToWmlConverterSettings htmlConverterSettings =
            HtmlToWmlConverter.GetDefaultSettings();

        private static readonly Regex detectEntityRegEx = new Regex("^&(?:#([0-9]+)|#x([0-9a-fA-F]+)|([0-9a-zA-Z]+));");

        private static readonly XElement softBreak = new XElement(W.r, new XElement(W.br));
        private static readonly XElement emptyRun = new XElement(W.r, new XElement(W.t));
        private static readonly XElement softTab = new XElement(W.r, new XElement(W.tab));

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
        internal static List<XElement> ProcessContentElement(
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
                return new List<XElement> { emptyRun };
            }

            // otherwise split the values if there are new line characters
            return ConvertTextToRunsWithMarkupSupport(values, part, templateError);
        }

        internal static List<XElement> ConvertTextToRunsWithMarkupSupport(
            string[] values,
            OpenXmlPart part,
            TemplateError templateError
        )
        {
            List<XElement> results = new List<XElement>();
            for (int i = 0; i < values.Length; i++)
            {
                string value = values[i];

                // empty and not the first element, this was a new line character and should be a soft break
                if (i > 0 && string.IsNullOrWhiteSpace(value))
                {
                    results.Add(softBreak);
                    continue;
                }

                // parse as XML
                XElement parsedElement = XElement.Parse($"<xhtml>{EscapeAmpersands(value)}</xhtml>");

                // check whether this is plain text and add runs if so
                if (parsedElement.IsPlainText())
                {
                    foreach (
                        var run in parsedElement
                            .Value.Replace("\r\n", "\n", StringComparison.OrdinalIgnoreCase)
                            .SplitAndKeep('\n')
                    )
                    {
                        if (run == "\n")
                            results.Add(softBreak);
                        else
                        {
                            foreach (var splitRun in run.SplitAndKeep('\t'))
                            {
                                if (splitRun == "\t")
                                    results.Add(softTab);
                                else
                                    results.Add(new XElement(W.r, new XElement(W.t, splitRun)));
                            }
                        }
                    }
                }
                else
                {
                    if (i > 0 && results.Last() != softBreak)
                    {
                        // if this is not the first element we are processing then add a soft break before
                        // only if our last run is not a soft-break itself!
                        results.Add(softBreak);
                    }

                    // otherwise we have XML let's process it
                    results.AddRange(
                        AddLineBreaks(
                            FlattenResults(
                                Transform(parsedElement, htmlConverterSettings, part, NextExpected.Run, true)
                            )
                        )
                    );
                }
            }

            if (results.Count == 0)
            {
                return new List<XElement> { emptyRun };
            }

            return results;
        }

        private static List<object> FlattenResults(object obj)
        {
            // flatten the returned content
            List<object> results = new List<object>();
            if (obj is IEnumerable)
            {
                results.AddRange(obj as IEnumerable<object>);
            }
            else
            {
                results.Add(obj);
            }

            return results;
        }

        private static List<XElement> AddLineBreaks(List<object> content)
        {
            List<XElement> result = new List<XElement>();
            for (int i = 0; i < content.Count; i++)
            {
                object obj = content[i];
                if (obj is XElement)
                {
                    // add a soft break between
                    if (i > 0)
                        result.Add(softBreak);

                    XElement element = obj as XElement;
                    IEnumerable<XElement> runs = element.DescendantsAndSelf(W.r);
                    if (runs != null && runs.Any())
                    {
                        foreach (var run in runs)
                        {
                            if (run.Parent != null && run.Parent.Name == W.hyperlink)
                            {
                                result.Add(run.Parent);
                            }
                            else
                            {
                                result.Add(run);
                            }
                        }
                    }
                }
            }

            return result;
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
