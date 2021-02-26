using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint
{
    internal abstract class SlidePartData<T> : IComparable<SlidePartData<T>>
    {
        public T Part { get; }
        private double ScaleFactor { get; }
        private string ShapeXml { get; }

        protected SlidePartData(T part, double scaleFactor)
        {
            Part = part;
            ScaleFactor = scaleFactor;
            ShapeXml = GetShapeDescriptor(part);
        }

        protected abstract string GetShapeDescriptor(T part);

        public virtual int CompareTo(SlidePartData<T> other)
        {
            if (ReferenceEquals(this, other)) return 0;
            if (ReferenceEquals(null, other)) return 1;
            return string.Compare(ShapeXml, other.ShapeXml, StringComparison.Ordinal);
        }
        
        protected string NormalizeXml(string xml)
        {
            var doc = XDocument.Parse(xml);
            CleanUpAttributes(doc.Root);
            ScaleShapes(doc, ScaleFactor);
            return doc.ToString();
        }
        
        private static readonly XNamespace s_relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XName[] s_noiseAttNames = {"smtClean", "dirty", "userDrawn", s_relNs + "id", s_relNs + "embed"};
        
        /// <summary>
        /// Remove OpenXml attributes that may occur on Layout/Master elements but does not affect rendering   
        /// </summary>
        private static void CleanUpAttributes(XElement element)
        {
            foreach (var attName in s_noiseAttNames)
            {
                element.Attribute(attName)?.Remove();
            }
            
            foreach (var descendant in element.Descendants()) 
            {
                CleanUpAttributes(descendant);
            }
        }

        public static void ScaleShapes(XDocument openXmlPart, double scale)
        {
            if (Math.Abs(scale - 1.0) < 1.0e-5)
                return;

            var shapeTree = openXmlPart.Descendants(P.spTree).FirstOrDefault();
            if (shapeTree is not null)
            {
                foreach (var transform in shapeTree.Descendants(A.xfrm))
                {
                    if (transform.Element(A.off) is {} offset)
                    {
                        Scale(offset, "x");
                        Scale(offset, "y");
                    }

                    if (transform.Element(A.ext) is {} extents)
                    {
                        Scale(extents, "cx");
                        Scale(extents, "cy");
                    }
                }

                foreach (var rPr in shapeTree.Descendants(A.rPr))
                {
                    Scale(rPr, "sz");
                }
            }

            void Scale(XElement element, string attrName)
            {
                var attr = element.Attribute(attrName);
                if (attr is null)
                    return;
                if (!long.TryParse(attr.Value, out var num))
                    return;

                var newNum = (long) (num * scale);
                attr.SetValue(newNum);
            }
        }
    }
    
    // This class is used to prevent duplication of layouts and handle content modification
    internal class SlideLayoutData : SlidePartData<SlideLayoutPart>
    {
        public SlideLayoutData(SlideLayoutPart slideLayout, double scaleFactor):base(slideLayout, scaleFactor) {}

        protected override string GetShapeDescriptor(SlideLayoutPart slideLayout) =>
            NormalizeXml(slideLayout.SlideLayout.CommonSlideData.ShapeTree.OuterXml);

    }
    
    // This class is used to prevent duplication of themes and handle content modification
    internal class ThemeData : SlidePartData<ThemePart>
    {
        public ThemeData(ThemePart themePart, double scaleFactor):base(themePart, scaleFactor) {}

        protected override string GetShapeDescriptor(ThemePart themePart) =>
            NormalizeXml(themePart.Theme.ThemeElements.OuterXml);
    }
    
    // This class is used to prevent duplication of masters and handle content modification
    internal class SlideMasterData : SlidePartData<SlideMasterPart>
    {
        public ThemeData ThemeData { get; }
        public List<SlideLayoutData> SlideLayoutList { get; }

        public SlideMasterData(SlideMasterPart slideMaster, double scaleFactor):base(slideMaster, scaleFactor)
        {
            ThemeData = new ThemeData(slideMaster.ThemePart, scaleFactor);
            SlideLayoutList = new List<SlideLayoutData>();
        }

        protected override string GetShapeDescriptor(SlideMasterPart slideMaster)
        {
            var sb = new StringBuilder();
            sb.Append(NormalizeXml(slideMaster.SlideMaster.CommonSlideData.ShapeTree.OuterXml));
            sb.Append(NormalizeXml(slideMaster.SlideMaster.ColorMap.OuterXml));
            return sb.ToString();
        }

        public override int CompareTo(SlidePartData<SlideMasterPart> other)
        {
            var res = base.CompareTo(other);
            if (res == 0 && other is SlideMasterData otherData)
                res = ThemeData.CompareTo(otherData.ThemeData);
            return res;
        }
    }
}
