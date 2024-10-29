using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    }
}
