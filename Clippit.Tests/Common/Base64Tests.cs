// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="Base64.ChunkBase64"/> (internal) and the
/// public <see cref="Base64.ConvertFromBase64"/> helper.
/// </summary>
public class Base64Tests
{
    private static readonly string NL = Environment.NewLine;

    // ── ChunkBase64: empty input ──────────────────────────────────────────

    [Test]
    public async Task B64_001_Empty_NoTrailingNewline_ReturnsEmpty()
    {
        var result = Base64.ChunkBase64("", appendTrailingNewline: false);
        await Assert.That(result).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task B64_002_Empty_TrailingNewline_ReturnsEmpty()
    {
        // An empty input produces no lines, so nothing to append a newline to.
        var result = Base64.ChunkBase64("", appendTrailingNewline: true);
        await Assert.That(result).IsEqualTo(string.Empty);
    }

    // ── ChunkBase64: input shorter than one line (76 chars) ──────────────

    [Test]
    public async Task B64_003_ShortInput_NoTrailingNewline_NothingAppended()
    {
        const string input = "abc";
        var result = Base64.ChunkBase64(input, appendTrailingNewline: false);
        await Assert.That(result).IsEqualTo("abc");
    }

    [Test]
    public async Task B64_004_ShortInput_TrailingNewline_NewlineAppended()
    {
        const string input = "abc";
        var result = Base64.ChunkBase64(input, appendTrailingNewline: true);
        await Assert.That(result).IsEqualTo("abc" + NL);
    }

    // ── ChunkBase64: input exactly 76 chars (one full line) ───────────────

    [Test]
    public async Task B64_005_ExactlyOneLine_NoTrailingNewline_NoNewline()
    {
        var input = new string('A', 76);
        var result = Base64.ChunkBase64(input, appendTrailingNewline: false);
        await Assert.That(result).IsEqualTo(input);
    }

    [Test]
    public async Task B64_006_ExactlyOneLine_TrailingNewline_NewlineAppended()
    {
        var input = new string('A', 76);
        var result = Base64.ChunkBase64(input, appendTrailingNewline: true);
        await Assert.That(result).IsEqualTo(input + NL);
    }

    // ── ChunkBase64: input spans exactly two full lines (152 chars) ───────

    [Test]
    public async Task B64_007_TwoFullLines_NoTrailingNewline_SeparatorBetween()
    {
        var line1 = new string('A', 76);
        var line2 = new string('B', 76);
        var input = line1 + line2;
        var result = Base64.ChunkBase64(input, appendTrailingNewline: false);
        await Assert.That(result).IsEqualTo(line1 + NL + line2);
    }

    [Test]
    public async Task B64_008_TwoFullLines_TrailingNewline_NewlineAfterBoth()
    {
        var line1 = new string('A', 76);
        var line2 = new string('B', 76);
        var input = line1 + line2;
        var result = Base64.ChunkBase64(input, appendTrailingNewline: true);
        await Assert.That(result).IsEqualTo(line1 + NL + line2 + NL);
    }

    // ── ChunkBase64: input 77 chars — one full line + 1 char remainder ────

    [Test]
    public async Task B64_009_OneLineAndRemainder_NoTrailingNewline_LastChunkHasNoNewline()
    {
        var line1 = new string('X', 76);
        const string remainder = "Z";
        var input = line1 + remainder;
        var result = Base64.ChunkBase64(input, appendTrailingNewline: false);
        await Assert.That(result).IsEqualTo(line1 + NL + remainder);
    }

    // ── ChunkBase64: line content is preserved verbatim ───────────────────

    [Test]
    public async Task B64_010_LineContentPreservedVerbatim()
    {
        // Use a realistic base64 alphabet string.
        const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
        // Build a 76-char string from the alphabet.
        var input = string.Concat(Enumerable.Repeat(chars, 2)).Substring(0, 76);

        var result = Base64.ChunkBase64(input, appendTrailingNewline: false);
        await Assert.That(result).IsEqualTo(input);
    }

    // ── ChunkBase64: three lines ──────────────────────────────────────────

    [Test]
    public async Task B64_011_ThreeFullLines_TwoSeparators()
    {
        var line1 = new string('A', 76);
        var line2 = new string('B', 76);
        var line3 = new string('C', 76);
        var input = line1 + line2 + line3;
        var result = Base64.ChunkBase64(input, appendTrailingNewline: false);
        await Assert.That(result).IsEqualTo(line1 + NL + line2 + NL + line3);
    }

    // ── ChunkBase64: round-trip via Convert.ToBase64String ────────────────

    [Test]
    public async Task B64_012_RoundTrip_Bytes_CanBeRecoveredAfterChunking()
    {
        var original = new byte[200];
        for (var i = 0; i < original.Length; i++)
            original[i] = (byte)(i % 256);

        var flat = Convert.ToBase64String(original);
        var chunked = Base64.ChunkBase64(flat, appendTrailingNewline: false);

        // Re-join by removing embedded newlines (same as ConvertFromBase64 does).
        var recovered = Convert.FromBase64String(chunked.Replace(NL, ""));
        await Assert.That(recovered).IsEquivalentTo(original);
    }

    // ── ConvertFromBase64: ignores \r\n line endings ──────────────────────

    [Test]
    public async Task B64_013_ConvertFromBase64_StripsCarriageReturnLineFeed()
    {
        var original = new byte[] { 0x01, 0x02, 0x03, 0xFF };
        var flat = Convert.ToBase64String(original);
        // Manually insert CRLF to simulate what ChunkBase64 + Windows NewLine produces.
        var withCrLf = flat[..2] + "\r\n" + flat[2..];

        var result = Base64.ConvertFromBase64("unused.bin", withCrLf);
        await Assert.That(result).IsEquivalentTo(original);
    }

    [Test]
    public async Task B64_014_ConvertFromBase64_FileNameIsIgnored_SameBytesAnyName()
    {
        var original = new byte[] { 10, 20, 30 };
        var b64 = Convert.ToBase64String(original);

        var r1 = Base64.ConvertFromBase64("foo.bin", b64);
        var r2 = Base64.ConvertFromBase64("bar.bin", b64);
        await Assert.That(r1).IsEquivalentTo(r2);
    }
}
