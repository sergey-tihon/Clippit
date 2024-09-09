// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Xml.Linq;

namespace Clippit
{
    public static partial class WmlComparer
    {
        private static bool False => false;
        private static bool SaveIntermediateFilesForDebugging => false;

        private static string NewLine => Environment.NewLine;

        private static XAttribute[] NamespaceAttributes =>
            new XAttribute[]
            {
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
            };

        private static XName[] RevElementsWithNoText => new[] { M.oMath, M.oMathPara, W.drawing };

        private static XName[] AttributesToTrimWhenCloning =>
            new[] { WP14.anchorId, WP14.editId, "ObjectID", "ShapeID", "id", "type" };

        private static int s_maxId;

        private static XName[] WordBreakElements =>
            new[]
            {
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
                W.endnoteReference,
            };

        private static XName[] AllowableRunChildren =>
            new[]
            {
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
                //W._object,
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
                W.instrText,
            };

        private static XName[] ElementsToThrowAway =>
            new[]
            {
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
                W.continuationSeparator,
            };

        private static XName[] ElementsToHaveSha1Hash =>
            new[] { W.p, W.tbl, W.tr, W.tc, W.drawing, W.pict, W.txbxContent };

        private static XName[] InvalidElements =>
            new[]
            {
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
                W.subDoc,
            };

        private static RecursionInfo[] RecursionElements =>
            new RecursionInfo[]
            {
                new() { ElementName = W.del, ChildElementPropertyNames = null },
                new() { ElementName = W.ins, ChildElementPropertyNames = null },
                new() { ElementName = W.tbl, ChildElementPropertyNames = new[] { W.tblPr, W.tblGrid, W.tblPrEx } },
                new() { ElementName = W.tr, ChildElementPropertyNames = new[] { W.trPr, W.tblPrEx } },
                new() { ElementName = W.tc, ChildElementPropertyNames = new[] { W.tcPr, W.tblPrEx } },
                new() { ElementName = W.pict, ChildElementPropertyNames = new[] { VML.shapetype } },
                new() { ElementName = VML.group, ChildElementPropertyNames = null },
                new() { ElementName = VML.shape, ChildElementPropertyNames = null },
                new() { ElementName = VML.rect, ChildElementPropertyNames = null },
                new() { ElementName = VML.textbox, ChildElementPropertyNames = null },
                new() { ElementName = O._lock, ChildElementPropertyNames = null },
                new() { ElementName = W.txbxContent, ChildElementPropertyNames = null },
                new() { ElementName = W10.wrap, ChildElementPropertyNames = null },
                new() { ElementName = W.sdt, ChildElementPropertyNames = new[] { W.sdtPr, W.sdtEndPr } },
                new() { ElementName = W.sdtContent, ChildElementPropertyNames = null },
                new() { ElementName = W.hyperlink, ChildElementPropertyNames = null },
                new() { ElementName = W.fldSimple, ChildElementPropertyNames = null },
                new() { ElementName = VML.shapetype, ChildElementPropertyNames = null },
                new() { ElementName = W.smartTag, ChildElementPropertyNames = new[] { W.smartTagPr } },
                new() { ElementName = W.ruby, ChildElementPropertyNames = new[] { W.rubyPr } },
            };

        private static XName[] ComparisonGroupingElements => new[] { W.p, W.tbl, W.tr, W.tc, W.txbxContent };
    }
}
