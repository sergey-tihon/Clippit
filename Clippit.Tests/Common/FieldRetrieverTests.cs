// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="FieldRetriever.ParseField"/>.
/// </summary>
public class FieldRetrieverTests
{
    // ── FR001: empty or whitespace-only input → empty FieldInfo ─────────────

    [Test]
    public async Task FR001_ParseField_EmptyString_ReturnsEmptyFieldInfo()
    {
        var result = FieldRetriever.ParseField(string.Empty);

        await Assert.That(result.FieldType).IsEqualTo(string.Empty);
        await Assert.That(result.Arguments).IsEmpty();
        await Assert.That(result.Switches).IsEmpty();
    }

    // ── FR002: unrecognised field type → empty FieldInfo ────────────────────

    [Test]
    [Arguments("DATE")]
    [Arguments("TIME")]
    [Arguments("PAGE")]
    [Arguments("AUTHOR")]
    [Arguments("MERGEFIELD Name")]
    public async Task FR002_ParseField_UnrecognisedFieldType_ReturnsEmptyFieldInfo(string field)
    {
        var result = FieldRetriever.ParseField(field);

        await Assert.That(result.FieldType).IsEqualTo(string.Empty);
        await Assert.That(result.Arguments).IsEmpty();
        await Assert.That(result.Switches).IsEmpty();
    }

    // ── FR003: HYPERLINK with no arguments ──────────────────────────────────

    [Test]
    public async Task FR003_ParseField_HyperlinkNoArguments_ReturnsSwitchesAndNoArguments()
    {
        var result = FieldRetriever.ParseField("HYPERLINK");

        await Assert.That(result.FieldType).IsEqualTo("HYPERLINK");
        await Assert.That(result.Arguments).IsEmpty();
        await Assert.That(result.Switches).IsEmpty();
    }

    // ── FR004: HYPERLINK with quoted URL ────────────────────────────────────

    [Test]
    public async Task FR004_ParseField_HyperlinkWithQuotedUrl_ParsesArgumentCorrectly()
    {
        var result = FieldRetriever.ParseField("HYPERLINK \"https://example.com\"");

        await Assert.That(result.FieldType).IsEqualTo("HYPERLINK");
        await Assert.That(result.Arguments).HasCount(1);
        await Assert.That(result.Arguments[0]).IsEqualTo("https://example.com");
        await Assert.That(result.Switches).IsEmpty();
    }

    // ── FR005: HYPERLINK with \\l (bookmark) switch ──────────────────────────

    [Test]
    public async Task FR005_ParseField_HyperlinkWithSwitch_ParsesSwitchCorrectly()
    {
        // ParseField classifies tokens by their first character: '\\' → switch, else → argument.
        // The switch parameter "Section1" does not start with '\\', so it appears in Arguments too.
        var result = FieldRetriever.ParseField(@"HYPERLINK ""https://example.com"" \l ""Section1""");

        await Assert.That(result.FieldType).IsEqualTo("HYPERLINK");
        // "https://example.com" and "Section1" are both non-switch tokens after FieldType
        await Assert.That(result.Arguments).HasCount(2);
        await Assert.That(result.Arguments[0]).IsEqualTo("https://example.com");
        await Assert.That(result.Arguments[1]).IsEqualTo("Section1");
        await Assert.That(result.Switches).HasCount(1);
        await Assert.That(result.Switches[0]).IsEqualTo(@"\l");
    }

    // ── FR006: REF with bookmark name and \\h switch ─────────────────────────

    [Test]
    public async Task FR006_ParseField_RefWithBookmarkAndSwitch_ParsesCorrectly()
    {
        var result = FieldRetriever.ParseField(@"REF MyBookmark \h");

        await Assert.That(result.FieldType).IsEqualTo("REF");
        await Assert.That(result.Arguments).HasCount(1);
        await Assert.That(result.Arguments[0]).IsEqualTo("MyBookmark");
        await Assert.That(result.Switches).HasCount(1);
        await Assert.That(result.Switches[0]).IsEqualTo(@"\h");
    }

    // ── FR007: SEQ with identifier ──────────────────────────────────────────

    [Test]
    public async Task FR007_ParseField_SeqWithIdentifier_ParsesCorrectly()
    {
        var result = FieldRetriever.ParseField("SEQ Figure");

        await Assert.That(result.FieldType).IsEqualTo("SEQ");
        await Assert.That(result.Arguments).HasCount(1);
        await Assert.That(result.Arguments[0]).IsEqualTo("Figure");
        await Assert.That(result.Switches).IsEmpty();
    }

    // ── FR008: STYLEREF with style name ─────────────────────────────────────

    [Test]
    public async Task FR008_ParseField_StylerefWithStyleName_ParsesCorrectly()
    {
        var result = FieldRetriever.ParseField("STYLEREF Heading1");

        await Assert.That(result.FieldType).IsEqualTo("STYLEREF");
        await Assert.That(result.Arguments).HasCount(1);
        await Assert.That(result.Arguments[0]).IsEqualTo("Heading1");
        await Assert.That(result.Switches).IsEmpty();
    }

    // ── FR009: field type comparison is case-insensitive ────────────────────

    [Test]
    [Arguments("hyperlink \"https://example.com\"", "HYPERLINK")]
    [Arguments("Hyperlink \"https://example.com\"", "HYPERLINK")]
    [Arguments("ref bookmark", "REF")]
    [Arguments("Ref bookmark", "REF")]
    [Arguments("seq Figure", "SEQ")]
    [Arguments("styleref Heading1", "STYLEREF")]
    public async Task FR009_ParseField_FieldTypeIsCaseInsensitive(string field, string expectedFieldType)
    {
        var result = FieldRetriever.ParseField(field);

        // ParseField accepts any case; the FieldType preserves the original casing from the token.
        await Assert.That(result.FieldType.Equals(expectedFieldType, StringComparison.OrdinalIgnoreCase)).IsTrue();
        await Assert.That(result.Arguments).IsNotEmpty();
    }

    // ── FR010: leading whitespace is tolerated ───────────────────────────────

    [Test]
    public async Task FR010_ParseField_LeadingWhitespace_ParsesFieldTypeCorrectly()
    {
        var result = FieldRetriever.ParseField("  REF MyBookmark");

        await Assert.That(result.FieldType).IsEqualTo("REF");
        await Assert.That(result.Arguments).HasCount(1);
        await Assert.That(result.Arguments[0]).IsEqualTo("MyBookmark");
    }
}
