using System;
using System.Collections.Generic;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint.Fluent;

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

    /// <summary>
    /// Normalize an XElement for shape comparison.
    /// Deep-clones <paramref name="element"/>, strips noise attributes, scales dimensions,
    /// and returns the canonical XML string. The original element is not modified.
    /// </summary>
    protected string NormalizeXml(XElement element)
    {
        var clone = new XElement(element);
        CleanUpAttributes(clone);
        ScaleShapes(clone, ScaleFactor);
        return clone.ToString();
    }

    public virtual int CompareTo(SlidePartData<T> other)
    {
        if (ReferenceEquals(this, other))
            return 0;
        if (other is null)
            return 1;
        return string.Compare(ShapeXml, other.ShapeXml, StringComparison.Ordinal);
    }

    private static readonly XNamespace s_relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
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

    private static readonly Dictionary<XName, string[]> s_resizableAttributes = new()
    {
        { A.off, ["x", "y"] }, // <a:off x="2054132" y="1665577"/>
        { A.ext, ["cx", "cy"] }, // <a:ext cx="2289267" cy="3074329"/>
        { A.chOff, ["x", "y"] }, // <a:chOff x="698501" y="1640632"/>
        { A.chExt, ["cx", "cy"] }, // <a:chExt cx="906462" cy="270006"/>
        { A.rPr, ["sz"] }, // <a:rPr lang="en-US" sz="700" b="1">
        { A.defRPr, ["sz"] }, // <a:defRPr sz="1350" kern="1200">
        { A.pPr, ["defTabSz"] }, // <a:pPr defTabSz="457119">
        { A.endParaRPr, ["sz"] }, // <a:endParaRPr lang="en-US" sz="2400" kern="0">
        { A.gridCol, ["w"] }, // <a:gridCol w="347223">
        { A.tr, ["h"] }, // <a:tr h="229849">
    };

    public static void ScaleShapes(XElement root, double scale)
    {
        if (Math.Abs(scale - 1.0) < 1.0e-5)
            return;

        foreach (var element in root.DescendantsAndSelf())
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

    /// <summary>Overload for callers that hold an XDocument (scales the root element).</summary>
    public static void ScaleShapes(XDocument openXmlPart, double scale) => ScaleShapes(openXmlPart.Root!, scale);
}

// This class is used to prevent duplication of layouts and handle content modification
internal class SlideLayoutData(SlideLayoutPart slideLayout, double scaleFactor)
    : SlidePartData<SlideLayoutPart>(slideLayout, scaleFactor)
{
    protected override string GetShapeDescriptor(SlideLayoutPart slideLayout)
    {
        var root = slideLayout.GetXDocument().Root!;
        var cSld = root.Element(P.cSld)!;
        var bg = cSld.Element(P.bg);
        return bg is not null
            ? string.Concat(NormalizeXml(cSld.Element(P.spTree)!), NormalizeXml(bg))
            : NormalizeXml(cSld.Element(P.spTree)!);
    }
}

// This class is used to prevent duplication of themes and handle content modification
internal class ThemeData(ThemePart themePart, double scaleFactor) : SlidePartData<ThemePart>(themePart, scaleFactor)
{
    protected override string GetShapeDescriptor(ThemePart themePart) =>
        NormalizeXml(themePart.GetXDocument().Root!.Element(A.themeElements)!);
}

// This class is used to prevent duplication of masters and handle content modification
internal class SlideMasterData(SlideMasterPart slideMaster, double scaleFactor)
    : SlidePartData<SlideMasterPart>(slideMaster, scaleFactor)
{
    public ThemeData ThemeData { get; } = new(slideMaster.ThemePart, scaleFactor);
    public Dictionary<SlideLayoutPart, SlideLayoutData> SlideLayouts { get; } = [];

    protected override string GetShapeDescriptor(SlideMasterPart slideMaster)
    {
        var root = slideMaster.GetXDocument().Root!;
        var cSld = root.Element(P.cSld)!;
        var bg = cSld.Element(P.bg);
        return bg is not null
            ? string.Concat(
                NormalizeXml(cSld.Element(P.spTree)!),
                NormalizeXml(bg),
                NormalizeXml(root.Element(P.clrMap)!)
            )
            : string.Concat(NormalizeXml(cSld.Element(P.spTree)!), NormalizeXml(root.Element(P.clrMap)!));
    }

    public override int CompareTo(SlidePartData<SlideMasterPart> other)
    {
        var res = base.CompareTo(other);
        if (res == 0 && other is SlideMasterData otherData)
            res = ThemeData.CompareTo(otherData.ThemeData);
        return res;
    }
}
