using System.Collections;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace Clippit.Word.Assembler
{
    internal static class XPathExtensions
    {
        internal static string[] EvaluateXPath(this XElement element, string xPath, bool optional)
        {
            object xPathSelectResult;
            try
            {
                // support some cells in the table may not have an xpath expression.
                if (string.IsNullOrWhiteSpace(xPath))
                {
                    return [];
                }

                xPathSelectResult = element.XPathEvaluate(xPath);
            }
            catch (XPathException e)
            {
                throw new XPathException("XPathException: " + e.Message, e);
            }

            if (xPathSelectResult is IEnumerable enumerable and not string)
            {
                var result = enumerable
                    .Cast<XObject>()
                    .Select(x =>
                        x switch
                        {
                            XElement xElement => xElement.Value,
                            XAttribute attribute => attribute.Value,
                            _ => throw new ArgumentException($"Unknown element type: {x.GetType().Name}"),
                        }
                    )
                    .ToArray();

                if (result.Length == 0 && !optional)
                    throw new XPathException($"XPath expression ({xPath}) returned no results");
                return result;
            }

            return new[] { xPathSelectResult.ToString() };
        }

        internal static string EvaluateXPathToString(this XElement element, string xPath, bool optional)
        {
            var selectedData = element.EvaluateXPath(xPath, true);

            return selectedData.Length switch
            {
                0 when optional => string.Empty,
                0 => throw new XPathException($"XPath expression ({xPath}) returned no results"),
                > 1 => throw new XPathException($"XPath expression ({xPath}) returned more than one node"),
                _ => selectedData.First(),
            };
        }

        internal static bool TryEvalueStringToByteArray(this XElement element, string pathOrXPath, out byte[] bytes)
        {
            bytes = [];

            try
            {
                var fileInfo = element.EvaluateStringToFileInfo(pathOrXPath);
                if (fileInfo != null)
                {
                    using (
                        var fs = new FileStream(fileInfo.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                    )
                    {
                        using (var ms = new MemoryStream())
                        {
                            fs.CopyTo(ms);
                            bytes = ms.ToArray();
                        }
                    }

                    return true;
                }
            }
            catch { }

            return false;
        }

        private static FileInfo EvaluateStringToFileInfo(this XElement element, string pathOrXPath)
        {
            object xPathSelectResult;
            try
            {
                xPathSelectResult = element.XPathEvaluate(pathOrXPath);
                if ((xPathSelectResult is IEnumerable) && !(xPathSelectResult is string))
                {
                    var selectedData = ((IEnumerable)xPathSelectResult).Cast<XObject>().SingleOrDefault();
                    if (selectedData != null)
                    {
                        if (selectedData.NodeType == XmlNodeType.Text)
                        {
                            XText text = selectedData as XText;
                            return new FileInfo(text.Value);
                        }
                        else if (selectedData.NodeType == XmlNodeType.Attribute)
                        {
                            XAttribute att = selectedData as XAttribute;
                            return new FileInfo(att.Value);
                        }
                        else if (selectedData.NodeType == XmlNodeType.Element)
                        {
                            // the element should have one child text node
                            XElement ele = selectedData as XElement;
                            XText text = ele.Nodes()
                                .Where(x => x.NodeType == XmlNodeType.Text)
                                .Select(x => x as XText)
                                .SingleOrDefault();

                            if (text != null)
                            {
                                return new FileInfo(text.Value);
                            }
                        }
                    }
                }
            }
            catch (XPathException) // suppress the xpath exception
            { }

            // check whether the xPath is actually just a file path
            try
            {
                return new FileInfo(pathOrXPath);
            }
            // supress exceptions that may occur if the path is actually xPath
            catch (ArgumentNullException) { }
            catch (NotSupportedException) { }
            catch (ArgumentException) { }

            return null;
        }
    }
}
