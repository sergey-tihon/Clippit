using System.Runtime.CompilerServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Drawing;

namespace Clippit.Word.Assembler
{
    internal static class XElementExtensions
    {
        internal static bool IsPlainText(this XElement element)
        {
            return element.Value == element.GetInnerXml();
        }

        internal static void MergeRunProperties(
            this XElement element,
            XElement paraRunProperties,
            XElement runRunProperties
        )
        {
            // merge run properties of paragraph properties
            if (element.Name == W.p && paraRunProperties != null)
            {
                XElement paraProps = element.Elements(W.pPr).FirstOrDefault();
                if (paraProps != null)
                {
                    XElement paraRunProps = paraProps.Elements(W.rPr).FirstOrDefault();
                    if (paraRunProps == null)
                    {
                        paraProps.Add(paraRunProperties);
                    }
                    else
                    {
                        paraRunProps.MergeOriginalRunProperties(paraRunProperties);
                    }
                }
            }

            // merge run properties of runs
            if (runRunProperties != null)
            {
                foreach (var run in element.DescendantsAndSelf(W.r))
                {
                    XElement runProps = run.Elements(W.rPr).FirstOrDefault();
                    if (runProps == null)
                    {
                        run.AddFirst(runRunProperties);
                    }
                    else
                    {
                        runProps.MergeOriginalRunProperties(runRunProperties);
                    }
                }
            }
        }

        private static void MergeOriginalRunProperties(this XElement runProps, XElement originalRunProps)
        {
            foreach (var prop in originalRunProps.Elements())
            {
                if (runProps.Element(prop.Name) == null)
                {
                    if (prop.Name == W.rStyle)
                    {
                        runProps.AddFirst(prop);
                    }
                    else
                    {
                        runProps.Add(prop);
                    }
                }
            }
        }

        private static string GetInnerXml(this XElement element)
        {
            string result = string.Empty;
            using (var reader = element.CreateReader())
            {
                reader.MoveToContent();
                result = reader.ReadInnerXml();
            }

            return System.Net.WebUtility.HtmlDecode(result);
        }

    }
}
