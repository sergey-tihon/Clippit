// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Word.Assembler;

namespace Clippit.Tests.Common;

public class PtUtilsTests
{
    // ── PtUtils.NormalizeDirName ────────────────────────────────────────────

    [Test]
    [Arguments("foo/bar", "foo/bar/")]
    [Arguments("foo/bar/", "foo/bar/")]
    [Arguments("foo\\bar", "foo/bar/")]
    [Arguments("foo\\bar\\", "foo/bar/")]
    [Arguments("a", "a/")]
    public async Task NormalizeDirName_AppendsSlashAndNormalizesBackslashes(string input, string expected)
    {
        await Assert.That(PtUtils.NormalizeDirName(input)).IsEqualTo(expected);
    }

    // ── PtUtils.MakeValidXml ────────────────────────────────────────────────

    [Test]
    public async Task MakeValidXml_StringWithNoControlChars_ReturnsUnchanged()
    {
        const string input = "Hello, World! ABC 123";
        await Assert.That(PtUtils.MakeValidXml(input)).IsEqualTo(input);
    }

    [Test]
    public async Task MakeValidXml_EmptyString_ReturnsEmpty()
    {
        await Assert.That(PtUtils.MakeValidXml(string.Empty)).IsEqualTo(string.Empty);
    }

    // Control chars (< 0x20) are encoded as _X_ where X is the unpadded uppercase
    // hex code-point: e.g. \x01 → "_1_", \x0A → "_A_", \x1F → "_1F_".
    [Test]
    [Arguments("\x01", "_1_")]
    [Arguments("\x0A", "_A_")]
    [Arguments("\x0D", "_D_")]
    [Arguments("\x1F", "_1F_")]
    public async Task MakeValidXml_SingleControlChar_IsEncoded(string input, string expected)
    {
        await Assert.That(PtUtils.MakeValidXml(input)).IsEqualTo(expected);
    }

    [Test]
    public async Task MakeValidXml_MixedString_OnlyControlCharsAreEncoded()
    {
        // \x01 (SOH) → "_1_", rest unchanged
        const string input = "abc\x01xyz";
        await Assert.That(PtUtils.MakeValidXml(input)).IsEqualTo("abc_1_xyz");
    }

    [Test]
    public async Task MakeValidXml_MultipleControlChars_AllEncoded()
    {
        // \x00 → "_0_", \x1F → "_1F_"
        const string input = "\x00\x1F";
        await Assert.That(PtUtils.MakeValidXml(input)).IsEqualTo("_0__1F_");
    }

    // ── StringExtensions.SplitAndKeep ──────────────────────────────────────

    [Test]
    public async Task SplitAndKeep_NoDelimitersFound_ReturnsSinglePart()
    {
        var result = "hello".SplitAndKeep(',');
        await Assert.That(result).IsEquivalentTo(["hello"]);
    }

    [Test]
    public async Task SplitAndKeep_EmptyString_ReturnsEmpty()
    {
        var result = string.Empty.SplitAndKeep(',');
        await Assert.That(result).IsEmpty();
    }

    [Test]
    public async Task SplitAndKeep_SingleDelimiter_SplitsIntoThreeParts()
    {
        var result = "a,b".SplitAndKeep(',');
        await Assert.That(result).IsEquivalentTo(["a", ",", "b"]);
    }

    [Test]
    public async Task SplitAndKeep_MultipleDelimiters_IncludesEachDelimiter()
    {
        var result = "a,b,c".SplitAndKeep(',');
        await Assert.That(result).IsEquivalentTo(["a", ",", "b", ",", "c"]);
    }

    [Test]
    public async Task SplitAndKeep_LeadingDelimiter_DelimiterIsFirstPart()
    {
        var result = ",a".SplitAndKeep(',');
        await Assert.That(result).IsEquivalentTo([",", "a"]);
    }

    [Test]
    public async Task SplitAndKeep_TrailingDelimiter_DelimiterIsLastPart()
    {
        var result = "a,".SplitAndKeep(',');
        await Assert.That(result).IsEquivalentTo(["a", ","]);
    }

    [Test]
    public async Task SplitAndKeep_OnlyDelimiter_ReturnsSingleDelimiterPart()
    {
        var result = ",".SplitAndKeep(',');
        await Assert.That(result).IsEquivalentTo([","]);
    }

    [Test]
    public async Task SplitAndKeep_MultipleDelimiterChars_EachIsKept()
    {
        var result = "a,b;c".SplitAndKeep(',', ';');
        await Assert.That(result).IsEquivalentTo(["a", ",", "b", ";", "c"]);
    }
}

public class PtExtensionsTests
{
    // ── ToBoolean ───────────────────────────────────────────────────────────

    [Test]
    public async Task ToBoolean_Null_ReturnsNull()
    {
        XAttribute? a = null;
        await Assert.That(a.ToBoolean()).IsNull();
    }

    [Test]
    [Arguments("1", true)]
    [Arguments("0", false)]
    [Arguments("true", true)]
    [Arguments("false", false)]
    [Arguments("on", true)]
    [Arguments("off", false)]
    [Arguments("True", true)]
    [Arguments("False", false)]
    public async Task ToBoolean_KnownValues_ReturnsExpected(string value, bool expected)
    {
        var elem = new XElement("e", new XAttribute("a", value));
        await Assert.That(elem.Attribute("a").ToBoolean()).IsEqualTo(expected);
    }

    // ── GetXPath ────────────────────────────────────────────────────────────

    [Test]
    public async Task GetXPath_Document_ReturnsDot()
    {
        var doc = new XDocument(new XElement("root"));
        await Assert.That(doc.GetXPath()).IsEqualTo(".");
    }

    [Test]
    public async Task GetXPath_RootElement_ReturnsSlashName()
    {
        var doc = new XDocument(new XElement("root"));
        await Assert.That(doc.Root!.GetXPath()).IsEqualTo("/root");
    }

    [Test]
    public async Task GetXPath_NestedElement_ReturnsFullPath()
    {
        var doc = new XDocument(new XElement("root", new XElement("child", new XElement("leaf"))));
        var leaf = doc.Root!.Element("child")!.Element("leaf")!;
        await Assert.That(leaf.GetXPath()).IsEqualTo("/root/child/leaf");
    }

    [Test]
    public async Task GetXPath_SiblingElements_IncludePredicate()
    {
        var doc = new XDocument(
            new XElement("root", new XElement("item", "a"), new XElement("item", "b"), new XElement("item", "c"))
        );
        var items = doc.Root!.Elements("item").ToList();
        await Assert.That(items[0].GetXPath()).IsEqualTo("/root/item[1]");
        await Assert.That(items[1].GetXPath()).IsEqualTo("/root/item[2]");
        await Assert.That(items[2].GetXPath()).IsEqualTo("/root/item[3]");
    }

    [Test]
    public async Task GetXPath_Attribute_ReturnsAtPath()
    {
        var doc = new XDocument(new XElement("root", new XAttribute("id", "42")));
        var attr = doc.Root!.Attribute("id")!;
        await Assert.That(attr.GetXPath()).IsEqualTo("/root/@id");
    }

    [Test]
    public async Task GetXPath_TextNode_ReturnsTextPath()
    {
        var doc = new XDocument(new XElement("root", new XElement("child", "hello")));
        var text = doc.Root!.Element("child")!.Nodes().OfType<XText>().First();
        await Assert.That(text.GetXPath()).IsEqualTo("/root/child/text()");
    }

    // ── GroupAdjacent ───────────────────────────────────────────────────────

    [Test]
    public async Task GroupAdjacent_EmptySequence_ReturnsEmpty()
    {
        var result = Enumerable.Empty<int>().GroupAdjacent(x => x).ToList();
        await Assert.That(result).IsEmpty();
    }

    [Test]
    public async Task GroupAdjacent_AllSameKey_ReturnsOneGroup()
    {
        var result = new[] { 1, 1, 1 }.GroupAdjacent(x => x).ToList();
        await Assert.That(result).HasCount(1);
        await Assert.That(result[0].Key).IsEqualTo(1);
        await Assert.That(result[0].ToList()).IsEquivalentTo([1, 1, 1]);
    }

    [Test]
    public async Task GroupAdjacent_AllDifferentKeys_ReturnsOneGroupPerElement()
    {
        var result = new[] { 1, 2, 3 }.GroupAdjacent(x => x).ToList();
        await Assert.That(result).HasCount(3);
    }

    [Test]
    public async Task GroupAdjacent_AdjacentDuplicates_GroupedTogether()
    {
        var result = new[] { 1, 1, 2, 2, 1, 1 }.GroupAdjacent(x => x).ToList();
        await Assert.That(result).HasCount(3);
        await Assert.That(result[0].Key).IsEqualTo(1);
        await Assert.That(result[1].Key).IsEqualTo(2);
        await Assert.That(result[2].Key).IsEqualTo(1);
        await Assert.That(result[0].Count()).IsEqualTo(2);
        await Assert.That(result[1].Count()).IsEqualTo(2);
        await Assert.That(result[2].Count()).IsEqualTo(2);
    }

    // ── StrCat ──────────────────────────────────────────────────────────────

    [Test]
    public async Task StrCat_EmptySequence_ReturnsEmpty()
    {
        var result = Enumerable.Empty<string>().StrCat("/");
        await Assert.That(result).IsEmpty();
    }

    [Test]
    public async Task StrCat_SingleElement_NoSeparatorAppended()
    {
        var result = new[] { "a" }.StrCat("/");
        await Assert.That(result).IsEqualTo("a/");
    }

    [Test]
    public async Task StrCat_MultipleElements_JoinedWithSeparator()
    {
        var result = new[] { "a", "b", "c" }.StrCat("/");
        await Assert.That(result).IsEqualTo("a/b/c/");
    }
}
