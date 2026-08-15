// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

/// <summary>
/// Unit tests for <see cref="ListItemTextGetter_tr_TR.GetListItemText"/>.
/// </summary>
public class ListItemTextTrTrTests
{
    // ── cardinalText — out-of-range guard ────────────────────────────────────

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    [Arguments(99999, "99999")]
    public async Task LTrTr001_CardinalText_OutOfRange_FallsBackToDecimal(int number, string expected)
    {
        var result = ListItemTextGetter_tr_TR.GetListItemText("tr-TR", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText — basic values ──────────────────────────────────────────

    [Test]
    [Arguments(1, "Bir")]
    [Arguments(2, "İki")]
    [Arguments(10, "On")]
    [Arguments(19, "Ondokuz")]
    [Arguments(20, "Yirmi")]
    [Arguments(21, "Yirmibir")]
    [Arguments(100, "Bir yüz")]
    [Arguments(101, "Bir yüz bir")]
    [Arguments(1000, "Bir bin")]
    [Arguments(1001, "Bir bin bir")]
    [Arguments(1100, "Bir bin bir yüz")]
    [Arguments(2000, "İki bin")]
    public async Task LTrTr002_CardinalText_TypicalValues_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_tr_TR.GetListItemText("tr-TR", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText — out-of-range guard ─────────────────────────────────────

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    [Arguments(99999, "99999")]
    public async Task LTrTr003_OrdinalText_OutOfRange_FallsBackToDecimal(int number, string expected)
    {
        var result = ListItemTextGetter_tr_TR.GetListItemText("tr-TR", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText — basic values ───────────────────────────────────────────

    [Test]
    [Arguments(1, "Birinci")]
    [Arguments(2, "İkinci")]
    [Arguments(10, "Onuncu")]
    [Arguments(19, "Ondokuzuncu")]
    [Arguments(20, "Yirminci")]
    [Arguments(100, "Yüzüncü")]
    [Arguments(101, "Yüzbirinci")]
    [Arguments(1000, "Bininci")]
    [Arguments(1001, "Binbirinci")]
    [Arguments(1100, "Binyüzüncü")]
    [Arguments(2000, "İkibininci")]
    public async Task LTrTr004_OrdinalText_TypicalValues_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_tr_TR.GetListItemText("tr-TR", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── decimal ──────────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "1")]
    [Arguments(42, "42")]
    public async Task LTrTr005_Decimal_ReturnsNumberAsString(int number, string expected)
    {
        var result = ListItemTextGetter_tr_TR.GetListItemText("tr-TR", number, "decimal");
        await Assert.That(result).IsEqualTo(expected);
    }
}
