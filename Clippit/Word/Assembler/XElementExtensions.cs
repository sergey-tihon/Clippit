﻿using System;
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
    }
}
