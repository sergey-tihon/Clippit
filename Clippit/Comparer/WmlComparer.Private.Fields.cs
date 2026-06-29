// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Frozen;
using System.Xml.Linq;

namespace Clippit
{
    public static partial class WmlComparer
    {
        private static bool False => false;
        private static bool SaveIntermediateFilesForDebugging => false;

        private static string NewLine => Environment.NewLine;

        // XAttribute instances in LINQ to XML are cloned when added to a new XElement if they
        // already have a parent, so this array can safely be shared across multiple XElement
        // constructions.
        private static readonly XAttribute[] NamespaceAttributes =
        [
            new(XNamespace.Xmlns + "wpc", WPC.wpc),
            new(XNamespace.Xmlns + "mc", MC.mc),
            new(XNamespace.Xmlns + "o", O.o),
            new(XNamespace.Xmlns + "r", R.r),
            new(XNamespace.Xmlns + "m", M.m),
            new(XNamespace.Xmlns + "v", VML.vml),
            new(XNamespace.Xmlns + "wp14", WP14.wp14),
            new(XNamespace.Xmlns + "wp", WP.wp),
            new(XNamespace.Xmlns + "w10", W10.w10),
            new(XNamespace.Xmlns + "w", W.w),
            new(XNamespace.Xmlns + "w14", W14.w14),
            new(XNamespace.Xmlns + "wpg", WPG.wpg),
            new(XNamespace.Xmlns + "wpi", WPI.wpi),
            new(XNamespace.Xmlns + "wne", WNE.wne),
            new(XNamespace.Xmlns + "wps", WPS.wps),
            new(MC.Ignorable, "w14 wp14"),
        ];

        // FrozenSet gives faster O(1) Contains() for static read-only membership tests;
        // static readonly avoids per-call allocation.
        private static readonly FrozenSet<XName> s_revElementsWithNoText = FrozenSet.Create<XName>(
            M.oMath,
            M.oMathPara,
            W.drawing
        );

        private static readonly FrozenSet<XName> s_attributesToTrimWhenCloning = FrozenSet.Create<XName>(
            WP14.anchorId,
            WP14.editId,
            "ObjectID",
            "ShapeID",
            "id",
            "type"
        );

        private static int s_maxId;

        private static readonly FrozenSet<XName> s_wordBreakElements = FrozenSet.Create<XName>(
            W.pPr,
            W.tab,
            W.br,
            W.continuationSeparator,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.drawing,
            W.pict,
            W.endnoteRef,
            W.footnoteRef,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W._object,
            W.ptab,
            W.separator,
            W.sym,
            W.yearLong,
            W.yearShort,
            M.oMathPara,
            M.oMath,
            W.footnoteReference,
            W.endnoteReference
        );

        private static readonly FrozenSet<XName> s_allowableRunChildren = FrozenSet.Create<XName>(
            W.br,
            W.drawing,
            W.cr,
            W.dayLong,
            W.dayShort,
            W.footnoteReference,
            W.endnoteReference,
            W.monthLong,
            W.monthShort,
            W.noBreakHyphen,
            W.pgNum,
            W.ptab,
            W.softHyphen,
            W.sym,
            W.tab,
            W.yearLong,
            W.yearShort,
            M.oMathPara,
            M.oMath,
            W.fldChar,
            W.instrText
        );

        private static readonly FrozenSet<XName> s_elementsToThrowAway = FrozenSet.Create<XName>(
            W.bookmarkStart,
            W.bookmarkEnd,
            W.commentRangeStart,
            W.commentRangeEnd,
            W.lastRenderedPageBreak,
            W.proofErr,
            W.tblPr,
            W.sectPr,
            W.permEnd,
            W.permStart,
            W.footnoteRef,
            W.endnoteRef,
            W.separator,
            W.continuationSeparator
        );

        private static readonly FrozenSet<XName> s_elementsToHaveSha1Hash = FrozenSet.Create<XName>(
            W.p,
            W.tbl,
            W.tr,
            W.tc,
            W.drawing,
            W.pict,
            W.txbxContent
        );

        private static readonly FrozenSet<XName> s_invalidElements = FrozenSet.Create<XName>(
            W.altChunk,
            W.customXml,
            W.customXmlDelRangeEnd,
            W.customXmlDelRangeStart,
            W.customXmlInsRangeEnd,
            W.customXmlInsRangeStart,
            W.customXmlMoveFromRangeEnd,
            W.customXmlMoveFromRangeStart,
            W.customXmlMoveToRangeEnd,
            W.customXmlMoveToRangeStart,
            W.moveFrom,
            W.moveFromRangeStart,
            W.moveFromRangeEnd,
            W.moveTo,
            W.moveToRangeStart,
            W.moveToRangeEnd,
            W.subDoc
        );

        // Array (not HashSet) since this is searched by FirstOrDefault with a predicate on ElementName.
        private static readonly RecursionInfo[] s_recursionElements =
        [
            new() { ElementName = W.del, ChildElementPropertyNames = null },
            new() { ElementName = W.ins, ChildElementPropertyNames = null },
            new() { ElementName = W.tbl, ChildElementPropertyNames = [W.tblPr, W.tblGrid, W.tblPrEx] },
            new() { ElementName = W.tr, ChildElementPropertyNames = [W.trPr, W.tblPrEx] },
            new() { ElementName = W.tc, ChildElementPropertyNames = [W.tcPr, W.tblPrEx] },
            new() { ElementName = W.pict, ChildElementPropertyNames = [VML.shapetype] },
            new() { ElementName = VML.group, ChildElementPropertyNames = null },
            new() { ElementName = VML.shape, ChildElementPropertyNames = null },
            new() { ElementName = VML.rect, ChildElementPropertyNames = null },
            new() { ElementName = VML.textbox, ChildElementPropertyNames = null },
            new() { ElementName = O._lock, ChildElementPropertyNames = null },
            new() { ElementName = W.txbxContent, ChildElementPropertyNames = null },
            new() { ElementName = W10.wrap, ChildElementPropertyNames = null },
            new() { ElementName = W.sdt, ChildElementPropertyNames = [W.sdtPr, W.sdtEndPr] },
            new() { ElementName = W.sdtContent, ChildElementPropertyNames = null },
            new() { ElementName = W.hyperlink, ChildElementPropertyNames = null },
            new() { ElementName = W.fldSimple, ChildElementPropertyNames = null },
            new() { ElementName = VML.shapetype, ChildElementPropertyNames = null },
            new() { ElementName = W.smartTag, ChildElementPropertyNames = [W.smartTagPr] },
            new() { ElementName = W.ruby, ChildElementPropertyNames = [W.rubyPr] },
        ];

        private static readonly FrozenSet<XName> s_comparisonGroupingElements = FrozenSet.Create<XName>(
            W.p,
            W.tbl,
            W.tr,
            W.tc,
            W.txbxContent
        );
    }
}
