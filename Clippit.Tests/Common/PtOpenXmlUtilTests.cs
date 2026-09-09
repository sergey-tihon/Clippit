// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for pure helper methods in <see cref="Clippit.PtOpenXmlUtil"/>: <c>XmlUtil.GetXmlSpaceAttribute</c>,
/// <c>WordprocessingMLUtil.GetBoolProp</c>, and <c>WordprocessingMLUtil.GetFontSize</c>. These are otherwise only
/// exercised indirectly through higher-level document processing (HtmlConverter, DocumentBuilder, etc.).
/// </summary>
public class PtOpenXmlUtilTests
{
    // ── XmlUtil.GetXmlSpaceAttribute(string) ────────────────────────────────

    [Test]
    public async Task GetXmlSpaceAttribute_String_EmptyString_ReturnsNull()
    {
        var attr = XmlUtil.GetXmlSpaceAttribute(string.Empty);
        await Assert.That(attr).IsNull();
    }

    [Test]
    public async Task GetXmlSpaceAttribute_String_NoLeadingOrTrailingSpace_ReturnsNull()
    {
        var attr = XmlUtil.GetXmlSpaceAttribute("hello");
        await Assert.That(attr).IsNull();
    }

    [Test]
    public async Task GetXmlSpaceAttribute_String_LeadingSpace_ReturnsPreserveAttribute()
    {
        var attr = XmlUtil.GetXmlSpaceAttribute(" hello");
        await Assert.That(attr).IsNotNull();
        await Assert.That(attr.Name).IsEqualTo(XNamespace.Xml + "space");
        await Assert.That(attr.Value).IsEqualTo("preserve");
    }

    [Test]
    public async Task GetXmlSpaceAttribute_String_TrailingSpace_ReturnsPreserveAttribute()
    {
        var attr = XmlUtil.GetXmlSpaceAttribute("hello ");
        await Assert.That(attr).IsNotNull();
        await Assert.That(attr.Value).IsEqualTo("preserve");
    }

    [Test]
    public async Task GetXmlSpaceAttribute_String_SingleSpaceChar_ReturnsPreserveAttribute()
    {
        var attr = XmlUtil.GetXmlSpaceAttribute(" ");
        await Assert.That(attr).IsNotNull();
        await Assert.That(attr.Value).IsEqualTo("preserve");
    }

    // ── XmlUtil.GetXmlSpaceAttribute(char) ───────────────────────────────────

    [Test]
    public async Task GetXmlSpaceAttribute_Char_Space_ReturnsPreserveAttribute()
    {
        var attr = XmlUtil.GetXmlSpaceAttribute(' ');
        await Assert.That(attr).IsNotNull();
        await Assert.That(attr.Name).IsEqualTo(XNamespace.Xml + "space");
        await Assert.That(attr.Value).IsEqualTo("preserve");
    }

    [Test]
    public async Task GetXmlSpaceAttribute_Char_NonSpace_ReturnsNull()
    {
        var attr = XmlUtil.GetXmlSpaceAttribute('a');
        await Assert.That(attr).IsNull();
    }

    // ── WordprocessingMLUtil.GetBoolProp ────────────────────────────────────

    [Test]
    public async Task GetBoolProp_ElementNotPresent_ReturnsFalse()
    {
        var rPr = new XElement(W.rPr);
        await Assert.That(WordprocessingMLUtil.GetBoolProp(rPr, W.b)).IsFalse();
    }

    [Test]
    public async Task GetBoolProp_ElementPresentNoValAttribute_ReturnsTrue()
    {
        var rPr = new XElement(W.rPr, new XElement(W.b));
        await Assert.That(WordprocessingMLUtil.GetBoolProp(rPr, W.b)).IsTrue();
    }

    [Test]
    [Arguments("1", true)]
    [Arguments("true", true)]
    [Arguments("0", false)]
    [Arguments("false", false)]
    [Arguments("TRUE", true)]
    [Arguments("FALSE", false)]
    public async Task GetBoolProp_ValAttribute_ParsesCT_OnOffSemantics(string val, bool expected)
    {
        var rPr = new XElement(W.rPr, new XElement(W.b, new XAttribute(W.val, val)));
        await Assert.That(WordprocessingMLUtil.GetBoolProp(rPr, W.b)).IsEqualTo(expected);
    }

    [Test]
    public async Task GetBoolProp_UnrecognizedValAttribute_ReturnsFalse()
    {
        var rPr = new XElement(W.rPr, new XElement(W.b, new XAttribute(W.val, "garbage")));
        await Assert.That(WordprocessingMLUtil.GetBoolProp(rPr, W.b)).IsFalse();
    }

    // ── WordprocessingMLUtil.GetFontSize ─────────────────────────────────────

    [Test]
    public async Task GetFontSize_ParagraphWithSz_ReturnsValue()
    {
        var p = new XElement(
            W.p,
            new XElement(W.pPr, new XElement(W.rPr, new XElement(W.sz, new XAttribute(W.val, "24"))))
        );
        var size = WordprocessingMLUtil.GetFontSize(p);
        await Assert.That(size).IsEqualTo(24m);
    }

    [Test]
    public async Task GetFontSize_ParagraphWithoutRPr_ReturnsNull()
    {
        var p = new XElement(W.p, new XElement(W.pPr));
        var size = WordprocessingMLUtil.GetFontSize(p);
        await Assert.That(size).IsNull();
    }

    [Test]
    public async Task GetFontSize_Run_ReturnsSzValue()
    {
        var r = new XElement(W.r, new XElement(W.rPr, new XElement(W.sz, new XAttribute(W.val, "36"))));
        var size = WordprocessingMLUtil.GetFontSize(r);
        await Assert.That(size).IsEqualTo(36m);
    }

    [Test]
    public async Task GetFontSize_RunWithBidiLanguageType_ReturnsSzCsValue()
    {
        var r = new XElement(
            W.r,
            new XAttribute(PtOpenXml.LanguageType, "bidi"),
            new XElement(
                W.rPr,
                new XElement(W.sz, new XAttribute(W.val, "20")),
                new XElement(W.szCs, new XAttribute(W.val, "28"))
            )
        );
        var size = WordprocessingMLUtil.GetFontSize(r);
        await Assert.That(size).IsEqualTo(28m);
    }

    [Test]
    public async Task GetFontSize_RunWithoutRPr_ReturnsNull()
    {
        var r = new XElement(W.r);
        var size = WordprocessingMLUtil.GetFontSize(r);
        await Assert.That(size).IsNull();
    }

    [Test]
    public async Task GetFontSize_ElementNotParagraphOrRun_ReturnsNull()
    {
        var e = new XElement(W.tbl);
        var size = WordprocessingMLUtil.GetFontSize(e);
        await Assert.That(size).IsNull();
    }

    [Test]
    public async Task GetFontSize_LanguageTypeOverload_NullRPr_ReturnsNull()
    {
        var size = WordprocessingMLUtil.GetFontSize("western", null);
        await Assert.That(size).IsNull();
    }
}
