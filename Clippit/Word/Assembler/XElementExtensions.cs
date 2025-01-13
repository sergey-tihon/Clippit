using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Clippit.Word.Assembler
{
    internal static class XElementExtensions
    {
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
                    XElement paraRunProps = element.Elements(W.rPr).FirstOrDefault();
                    if (paraRunProps == null)
                    {
                        paraProps.Add(paraRunProperties);
                    }
                    else
                    {
                        foreach (var prop in paraRunProperties.Elements())
                        {
                            if (paraRunProps.Element(prop.Name) == null)
                            {
                                paraRunProps.Add(prop);
                            }
                        }
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
                        foreach (var prop in runRunProperties.Elements())
                        {
                            if (runProps.Element(prop.Name) == null)
                            {
                                runProps.Add(prop);
                            }
                        }
                    }
                }
            }
        }
    }
}
