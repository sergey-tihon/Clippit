using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint
{
    internal abstract class SlidePartData<T> : IComparable<SlidePartData<T>>
    {
        public T Part { get; }
        private string ShapeXml { get; }

        protected SlidePartData(T part)
        {
            Part = part;
            ShapeXml = GetShapeDescriptor(part);
        }

        protected abstract string GetShapeDescriptor(T part);

        public virtual int CompareTo(SlidePartData<T> other)
        {
            if (ReferenceEquals(this, other)) return 0;
            if (ReferenceEquals(null, other)) return 1;
            return string.Compare(ShapeXml, other.ShapeXml, StringComparison.Ordinal);
        }
        
        protected static string NormalizeXml(string xml)
        {
            var doc = XDocument.Parse(xml);
            CleanUpAttributes(doc.Root);
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

    }
    
    // This class is used to prevent duplication of layouts and handle content modification
    internal class SlideLayoutData : SlidePartData<SlideLayoutPart>
    {
        public SlideLayoutData(SlideLayoutPart slideLayout):base(slideLayout) {}

        protected override string GetShapeDescriptor(SlideLayoutPart slideLayout) =>
            NormalizeXml(slideLayout.SlideLayout.CommonSlideData.ShapeTree.OuterXml);

    }
    
    // This class is used to prevent duplication of themes and handle content modification
    internal class ThemeData : SlidePartData<ThemePart>
    {
        public ThemeData(ThemePart themePart):base(themePart) {}

        protected override string GetShapeDescriptor(ThemePart themePart) =>
            NormalizeXml(themePart.Theme.ThemeElements.OuterXml);
    }
    
    // This class is used to prevent duplication of masters and handle content modification
    internal class SlideMasterData : SlidePartData<SlideMasterPart>
    {
        public ThemeData ThemeData { get; }
        public List<SlideLayoutData> SlideLayoutList { get; }

        public SlideMasterData(SlideMasterPart slideMaster):base(slideMaster)
        {
            ThemeData = new ThemeData(slideMaster.ThemePart);
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
