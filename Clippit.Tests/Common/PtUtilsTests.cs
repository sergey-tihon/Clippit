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
    public async Task StrCat_SingleElement_TrailingSeparatorAppended()
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

    // ── StringConcatenate ──────────────────────────────────────────────────

    [Test]
    public async Task StringConcatenate_StringOverload_JoinsAll()
    {
        var result = new[] { "foo", "bar", "baz" }.StringConcatenate();
        await Assert.That(result).IsEqualTo("foobarbaz");
    }

    [Test]
    public async Task StringConcatenate_Empty_ReturnsEmpty()
    {
        var result = Enumerable.Empty<string>().StringConcatenate();
        await Assert.That(result).IsEmpty();
    }

    [Test]
    public async Task StringConcatenate_WithProjection_AppliesProjection()
    {
        var result = new[] { 1, 2, 3 }.StringConcatenate(n => n.ToString());
        await Assert.That(result).IsEqualTo("123");
    }

    // ── PtZip ─────────────────────────────────────────────────────────────

    [Test]
    public async Task PtZip_EqualLengthSequences_ZipsAll()
    {
        var result = new[] { 1, 2, 3 }.PtZip(new[] { "a", "b", "c" }, (n, s) => $"{n}{s}").ToList();
        await Assert.That(result).IsEquivalentTo(["1a", "2b", "3c"]);
    }

    [Test]
    public async Task PtZip_FirstShorter_StopsAtFirst()
    {
        var result = new[] { 1, 2 }.PtZip(new[] { "a", "b", "c" }, (n, s) => n + s).ToList();
        await Assert.That(result).HasCount(2);
    }

    [Test]
    public async Task PtZip_SecondShorter_StopsAtSecond()
    {
        var result = new[] { 1, 2, 3 }.PtZip(new[] { "a" }, (n, s) => n + s).ToList();
        await Assert.That(result).HasCount(1);
    }

    // ── SkipLast ──────────────────────────────────────────────────────────

    [Test]
    public async Task SkipLast_SkipZero_ReturnsAll()
    {
        var result = new[] { 1, 2, 3 }.SkipLast(0).ToList();
        await Assert.That(result).IsEquivalentTo([1, 2, 3]);
    }

    [Test]
    public async Task SkipLast_SkipTwo_ReturnsFirstElement()
    {
        var result = new[] { 1, 2, 3 }.SkipLast(2).ToList();
        await Assert.That(result).IsEquivalentTo([1]);
    }

    [Test]
    public async Task SkipLast_SkipMoreThanLength_ReturnsEmpty()
    {
        var result = new[] { 1, 2 }.SkipLast(5).ToList();
        await Assert.That(result).IsEmpty();
    }

    // ── SequenceAt ────────────────────────────────────────────────────────

    [Test]
    public async Task SequenceAt_FromStart_ReturnsAll()
    {
        var arr = new[] { 10, 20, 30 };
        var result = arr.SequenceAt(0).ToList();
        await Assert.That(result).IsEquivalentTo([10, 20, 30]);
    }

    [Test]
    public async Task SequenceAt_FromMiddle_ReturnsTail()
    {
        var arr = new[] { 10, 20, 30 };
        var result = arr.SequenceAt(1).ToList();
        await Assert.That(result).IsEquivalentTo([20, 30]);
    }

    [Test]
    public async Task SequenceAt_PastEnd_ReturnsEmpty()
    {
        var arr = new[] { 10, 20 };
        var result = arr.SequenceAt(5).ToList();
        await Assert.That(result).IsEmpty();
    }

    // ── Rollup ────────────────────────────────────────────────────────────

    [Test]
    public async Task Rollup_Empty_ReturnsEmpty()
    {
        var result = Enumerable.Empty<int>().Rollup(0, (x, acc) => acc + x).ToList();
        await Assert.That(result).IsEmpty();
    }

    [Test]
    public async Task Rollup_RunningSum_ProducesAccumulatedValues()
    {
        var result = new[] { 1, 2, 3, 4 }.Rollup(0, (x, acc) => acc + x).ToList();
        await Assert.That(result).IsEquivalentTo([1, 3, 6, 10]);
    }

    [Test]
    public async Task Rollup_WithIndex_IndexPassedCorrectly()
    {
        var result = new[] { "a", "b", "c" }.Rollup("", (x, acc, i) => $"{i}:{x}").ToList();
        await Assert.That(result).IsEquivalentTo(["0:a", "1:b", "2:c"]);
    }

    // ── DescendantsTrimmed ─────────────────────────────────────────────────

    [Test]
    public async Task DescendantsTrimmed_ByName_StopsAtTrimElement()
    {
        var root = XElement.Parse("<root><a><b><c/></b></a><b><d/></b></root>");
        XName trimName = "b";
        var result = root.DescendantsTrimmed(trimName).Select(e => e.Name.LocalName).ToList();
        // Expects: a, b (under a — trimmed, no c), b (child of root — trimmed, no d)
        await Assert.That(result).IsEquivalentTo(["a", "b", "b"]);
    }

    [Test]
    public async Task DescendantsTrimmed_ByPredicate_StopsWhenPredicateTrue()
    {
        var root = XElement.Parse("<root><keep><stop><deep/></stop></keep></root>");
        var result = root.DescendantsTrimmed(e => e.Name.LocalName == "stop").Select(e => e.Name.LocalName).ToList();
        await Assert.That(result).IsEquivalentTo(["keep", "stop"]);
    }

    // ── SiblingsBeforeSelfReverseDocumentOrder ─────────────────────────────

    [Test]
    public async Task SiblingsBeforeSelfReverseDocumentOrder_ReturnsOlderSiblingsInReverse()
    {
        var root = XElement.Parse("<root><a/><b/><c/><d/></root>");
        var d = root.Element("d");
        var result = d!.SiblingsBeforeSelfReverseDocumentOrder().Select(e => e.Name.LocalName).ToList();
        await Assert.That(result).IsEquivalentTo(["c", "b", "a"]);
    }

    [Test]
    public async Task SiblingsBeforeSelfReverseDocumentOrder_FirstSibling_ReturnsEmpty()
    {
        var root = XElement.Parse("<root><a/><b/></root>");
        var a = root.Element("a");
        var result = a!.SiblingsBeforeSelfReverseDocumentOrder().ToList();
        await Assert.That(result).IsEmpty();
    }

    // ── DescendantsBeforeSelfReverseDocumentOrder ──────────────────────────

    [Test]
    public async Task DescendantsBeforeSelfReverseDocumentOrder_ReturnsDescendantsInReverse()
    {
        var root = XElement.Parse("<root><a><b/></a><c/></root>");
        var c = root.Element("c");
        var result = c!.DescendantsBeforeSelfReverseDocumentOrder().Select(e => e.Name.LocalName).ToList();
        // In document order: root, a, b, c — so before c in reverse: b, a
        await Assert.That(result).IsEquivalentTo(["b", "a"]);
    }

    // ── ToStringNewLineOnAttributes ────────────────────────────────────────

    [Test]
    public async Task ToStringNewLineOnAttributes_ElementWithAttributes_EachAttributeOnOwnLine()
    {
        var el = new XElement("item", new XAttribute("id", "1"), new XAttribute("name", "foo"));
        var result = el.ToStringNewLineOnAttributes();
        await Assert.That(result).Contains("id=\"1\"");
        await Assert.That(result).Contains("name=\"foo\"");
        // Each attribute should be on its own line
        var lines = result.Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0).ToArray();
        await Assert.That(lines.Length).IsGreaterThan(1);
    }
}
