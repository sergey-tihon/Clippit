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
            if (ReferenceEquals(this, other))
                return 0;
            if (other is null)
                return 1;
            return string.Compare(ShapeXml, other.ShapeXml, StringComparison.Ordinal);
        }

        protected string NormalizeXml(string xml)
        {
            var doc = XDocument.Parse(xml);
            CleanUpAttributes(doc.Root);
            ScaleShapes(doc, ScaleFactor);
            return doc.ToString();
        }

        private static readonly XNamespace s_relNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XName[] s_noiseAttNames =
        [
            "smtClean",
            "dirty",
            "userDrawn",
            s_relNs + "id",
            s_relNs + "embed",
        ];

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

        private static readonly Dictionary<XName, string[]> s_resizableAttributes =
            new()
            {
                { A.off, new[] { "x", "y" } }, // <a:off x="2054132" y="1665577"/>
                { A.ext, new[] { "cx", "cy" } }, // <a:ext cx="2289267" cy="3074329"/>
                { A.chOff, new[] { "x", "y" } }, // <a:chOff x="698501" y="1640632"/>
                { A.chExt, new[] { "cx", "cy" } }, // <a:chExt cx="906462" cy="270006"/>
                { A.rPr, new[] { "sz" } }, // <a:rPr lang="en-US" sz="700" b="1">
                { A.defRPr, new[] { "sz" } }, // <a:defRPr sz="1350" kern="1200">
                { A.pPr, new[] { "defTabSz" } }, // <a:pPr defTabSz="457119">
                { A.endParaRPr, new[] { "sz" } }, // <a:endParaRPr lang="en-US" sz="2400" kern="0">
                { A.gridCol, new[] { "w" } }, // <a:gridCol w="347223">
                { A.tr, new[] { "h" } }, // <a:tr h="229849">
            };

        public static void ScaleShapes(XDocument openXmlPart, double scale)
        {
            if (Math.Abs(scale - 1.0) < 1.0e-5)
                return;

            foreach (var element in openXmlPart.Descendants())
            {
                if (!s_resizableAttributes.TryGetValue(element.Name, out var attrNames))
                    continue;

                foreach (var attrName in attrNames)
                {
                    var attr = element.Attribute(attrName);
                    if (attr is null)
                        continue;
                    if (!long.TryParse(attr.Value, out var num))
                        continue;

                    var newNum = (long)(num * scale);
                    attr.SetValue(newNum);
                }
            }
        }
    }

    // This class is used to prevent duplication of layouts and handle content modification
    internal class SlideLayoutData(SlideLayoutPart slideLayout, double scaleFactor)
        : SlidePartData<SlideLayoutPart>(slideLayout, scaleFactor)
    {
        protected override string GetShapeDescriptor(SlideLayoutPart slideLayout)
        {
            var sb = new StringBuilder();
            var cSld = slideLayout.SlideLayout.CommonSlideData;
            sb.Append(NormalizeXml(cSld.ShapeTree.OuterXml));
            if (cSld.Background is not null)
                sb.Append(NormalizeXml(cSld.Background.OuterXml));
            return sb.ToString();
        }
    }

    // This class is used to prevent duplication of themes and handle content modification
    internal class ThemeData(ThemePart themePart, double scaleFactor) : SlidePartData<ThemePart>(themePart, scaleFactor)
    {
        protected override string GetShapeDescriptor(ThemePart themePart) =>
            NormalizeXml(themePart.Theme.ThemeElements.OuterXml);
    }

    // This class is used to prevent duplication of masters and handle content modification
    internal class SlideMasterData(SlideMasterPart slideMaster, double scaleFactor)
        : SlidePartData<SlideMasterPart>(slideMaster, scaleFactor)
    {
        public ThemeData ThemeData { get; } = new ThemeData(slideMaster.ThemePart, scaleFactor);
        public List<SlideLayoutData> SlideLayoutList { get; } = [];

        protected override string GetShapeDescriptor(SlideMasterPart slideMaster)
        {
            var sb = new StringBuilder();
            var cSld = slideMaster.SlideMaster.CommonSlideData;
            sb.Append(NormalizeXml(cSld.ShapeTree.OuterXml));
            if (cSld.Background is not null)
                sb.Append(NormalizeXml(cSld.Background.OuterXml));

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
