// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit.Html;

namespace Clippit.Tests.Html;

/// <summary>
/// Unit tests for the pure helper members of HtmlToWmlConverter.cs (CleanUpCss, and the
/// Emu/TPoint/Twip value-type conversions) that do not require a full HTML-to-WML conversion pipeline.
/// </summary>
public class HtmlToWmlConverterUnitTests
{
    [Test]
    public async Task CleanUpCss_NullInput_ReturnsEmptyString()
    {
#pragma warning disable CS8625 // Cannot convert null literal to non-nullable reference type.
        var result = HtmlToWmlConverter.CleanUpCss(null);
#pragma warning restore CS8625 // Cannot convert null literal to non-nullable reference type.
        await Assert.That(result).IsEqualTo("");
    }

    [Test]
    public async Task CleanUpCss_EmptyInput_ReturnsSingleNewline()
    {
        // Splitting an empty string yields a single empty line, which is kept
        // (it isn't a comment marker) and has a trailing newline appended.
        var result = HtmlToWmlConverter.CleanUpCss("");
        await Assert.That(result).IsEqualTo(Environment.NewLine);
    }

    [Test]
    [Arguments("//")]
    [Arguments("////")]
    [Arguments("<!--")]
    [Arguments("&lt;!--")]
    [Arguments("-->")]
    [Arguments("--&gt;")]
    public async Task CleanUpCss_RemovesExactCommentMarkerLines(string commentLine)
    {
        var css = $"body {{ color: red; }}\n{commentLine}\np {{ color: blue; }}";
        var result = HtmlToWmlConverter.CleanUpCss(css);

        await Assert.That(result).Contains("body { color: red; }");
        await Assert.That(result).Contains("p { color: blue; }");

        // The marker line itself (once trimmed) must be filtered out, so it should not appear
        // as its own line in the output.
        var lines = result.Split(Environment.NewLine, StringSplitOptions.RemoveEmptyEntries);
        await Assert.That(lines).DoesNotContain(commentLine);
    }

    [Test]
    public async Task CleanUpCss_KeepsLineThatOnlyStartsWithCommentMarker()
    {
        // The filter only removes lines that are an exact trimmed match for a marker,
        // so a line merely starting with "//" (not exactly "//") should be preserved.
        var css = "// this is a comment, not just the marker\nbody { color: red; }";
        var result = HtmlToWmlConverter.CleanUpCss(css);

        await Assert.That(result).Contains("// this is a comment, not just the marker");
        await Assert.That(result).Contains("body { color: red; }");
    }

    [Test]
    public async Task CleanUpCss_TrimsLeadingAndTrailingWhitespace()
    {
        var css = "  \r\n  body { color: red; }  \r\n  ";
        var result = HtmlToWmlConverter.CleanUpCss(css);

        await Assert.That(result.TrimEnd()).IsEqualTo("body { color: red; }");
    }

    [Test]
    public async Task CleanUpCss_HandlesMixedLineEndings()
    {
        var css = "line1\r\nline2\rline3\nline4";
        var result = HtmlToWmlConverter.CleanUpCss(css);

        await Assert.That(result).Contains("line1");
        await Assert.That(result).Contains("line2");
        await Assert.That(result).Contains("line3");
        await Assert.That(result).Contains("line4");
    }

    [Test]
    public async Task Emu_TwipsToEmus_OneInch_Returns914400()
    {
        // 1 inch = 1440 twips = 914400 EMUs
        Emu emu = Emu.TwipsToEmus(1440);
        long value = emu;
        await Assert.That(value).IsEqualTo(914400L);
    }

    [Test]
    public async Task Emu_TwipsToEmus_Zero_ReturnsZero()
    {
        Emu emu = Emu.TwipsToEmus(0);
        long value = emu;
        await Assert.That(value).IsEqualTo(0L);
    }

    [Test]
    public async Task Emu_PointsToEmus_OneInch_Returns914400()
    {
        // 1 inch = 72 points = 914400 EMUs
        Emu emu = Emu.PointsToEmus(72);
        long value = emu;
        await Assert.That(value).IsEqualTo(914400L);
    }

    [Test]
    public async Task Emu_ImplicitConversions_RoundTripThroughLong()
    {
        Emu emu = 12345L;
        long value = emu;
        await Assert.That(value).IsEqualTo(12345L);
    }

    [Test]
    public async Task Emu_ToString_Throws()
    {
        Emu emu = 100L;
        await Assert.That(() => emu.ToString()).Throws<OpenXmlPowerToolsException>();
    }

    [Test]
    public async Task TPoint_ImplicitConversions_RoundTripThroughDouble()
    {
        TPoint point = 3.14;
        double value = point;
        await Assert.That(value).IsEqualTo(3.14);
    }

    [Test]
    public async Task TPoint_ToString_Throws()
    {
        TPoint point = 1.0;
        await Assert.That(() => point.ToString()).Throws<OpenXmlPowerToolsException>();
    }

    [Test]
    public async Task Twip_ImplicitConversionFromLong_RoundTrips()
    {
        Twip twip = 1440L;
        long value = twip;
        await Assert.That(value).IsEqualTo(1440L);
    }

    [Test]
    public async Task Twip_ImplicitConversionFromDouble_TruncatesToLong()
    {
        Twip twip = 1440.9;
        long value = twip;
        await Assert.That(value).IsEqualTo(1440L);
    }

    [Test]
    public async Task Twip_ToString_Throws()
    {
        Twip twip = 100L;
        await Assert.That(() => twip.ToString()).Throws<OpenXmlPowerToolsException>();
    }
}
