// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

public class ListItemTextDeDETests
{
    // ── Out-of-range guards ───────────────────────────────────────────────────

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    public async Task LDE000_OutOfRange_FallsBackToDecimal_CardinalText(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    public async Task LDE000b_OutOfRange_FallsBackToDecimal_OrdinalText(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0)]
    [Arguments(-1)]
    [Arguments(20000)]
    public async Task LDE000c_UnsupportedFormat_ReturnsNull(int number)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "unsupportedFormat");
        await Assert.That(result).IsNull();
    }

    // ── cardinalText: 1–19 ───────────────────────────────────────────────────

    [Test]
    [Arguments(1, "Ein")]
    [Arguments(2, "Zwei")]
    [Arguments(3, "Drei")]
    [Arguments(7, "Sieben")]
    [Arguments(11, "Elf")]
    [Arguments(12, "Zwölf")]
    [Arguments(19, "Neunzehn")]
    public async Task LDE001_CardinalText_OneThroughNineteen(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText: 20–99 ──────────────────────────────────────────────────

    [Test]
    [Arguments(20, "Zwanzig")]
    [Arguments(21, "Einundzwanzig")]
    [Arguments(30, "Dreißig")]
    [Arguments(42, "Zweiundvierzig")]
    [Arguments(55, "Fünfundfünfzig")]
    [Arguments(99, "Neunundneunzig")]
    public async Task LDE002_CardinalText_TwentyThroughNinetyNine(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText: hundreds ───────────────────────────────────────────────

    [Test]
    [Arguments(100, "Hundert")]
    [Arguments(200, "Zweihundert")]
    [Arguments(300, "Dreihundert")]
    [Arguments(115, "Hundertfünfzehn")]
    [Arguments(999, "Neunhundertneunundneunzig")]
    public async Task LDE003_CardinalText_Hundreds(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText: thousands ──────────────────────────────────────────────

    [Test]
    [Arguments(1000, "Tausend")]
    [Arguments(2000, "Zweitausend")]
    [Arguments(5001, "Fünftausendein")]
    [Arguments(19999, "Neunzehntausendneunhundertneunundneunzig")]
    public async Task LDE004_CardinalText_Thousands(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText: 1–19 ────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "Erste")]
    [Arguments(2, "Zweite")]
    [Arguments(3, "Dritte")]
    [Arguments(7, "Siebte")]
    [Arguments(8, "Achte")]
    [Arguments(12, "Zwölfte")]
    [Arguments(19, "Neunzehnte")]
    public async Task LDE010_OrdinalText_OneThroughNineteen(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText: 20–99 ───────────────────────────────────────────────────

    [Test]
    [Arguments(20, "Zwanzigste")]
    [Arguments(21, "Einundzwanzigste")]
    [Arguments(30, "Dreißigste")]
    [Arguments(42, "Zweiundvierzigste")]
    [Arguments(99, "Neunundneunzigste")]
    public async Task LDE011_OrdinalText_TwentyThroughNinetyNine(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText: hundreds ────────────────────────────────────────────────

    [Test]
    [Arguments(100, "Hundertste")]
    [Arguments(101, "Hunderterste")]
    [Arguments(115, "Hundertfünfzehnte")]
    [Arguments(120, "Hundertzwanzigste")]
    [Arguments(121, "Hunderteinundzwanzigste")]
    public async Task LDE012_OrdinalText_Hundreds(int number, string expected)
    {
        var result = ListItemTextGetter_de_DE.GetListItemText("de-DE", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }
}
