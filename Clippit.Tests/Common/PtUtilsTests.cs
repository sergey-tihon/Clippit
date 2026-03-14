// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

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
