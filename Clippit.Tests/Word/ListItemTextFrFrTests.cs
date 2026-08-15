// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

/// <summary>
/// Unit tests for <see cref="ListItemTextGetter_fr_FR.GetListItemText"/>.
/// </summary>
public class ListItemTextFrFrTests
{
    // ── cardinalText — ones / teens ──────────────────────────────────────────

    [Test]
    [Arguments(1, "Un")]
    [Arguments(2, "Deux")]
    [Arguments(9, "Neuf")]
    [Arguments(10, "Dix")]
    [Arguments(11, "Onze")]
    [Arguments(16, "Seize")]
    [Arguments(19, "Dix-neuf")]
    public async Task LFrFr001_CardinalText_OnesToNineteen_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_fr_FR.GetListItemText("fr-FR", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText — tens ──────────────────────────────────────────────────

    [Test]
    [Arguments(20, "Vingt")]
    [Arguments(21, "Vingt et un")]
    [Arguments(22, "Vingt-deux")]
    [Arguments(30, "Trente")]
    [Arguments(31, "Trente et un")]
    [Arguments(50, "Cinquante")]
    [Arguments(60, "Soixante")]
    [Arguments(70, "Soixante-dix")]
    [Arguments(71, "Soixante et onze")]
    [Arguments(72, "Soixante-douze")]
    [Arguments(79, "Soixante-dix-neuf")]
    [Arguments(80, "Quatre-vingts")]
    [Arguments(81, "Quatre-vingt-un")]
    [Arguments(90, "Quatre-vingt-dix")]
    [Arguments(91, "Quatre-vingt-onze")]
    [Arguments(99, "Quatre-vingt-dix-neuf")]
    public async Task LFrFr002_CardinalText_TwentyToNinetyNine_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_fr_FR.GetListItemText("fr-FR", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText — hundreds ──────────────────────────────────────────────

    [Test]
    [Arguments(100, "Cent")]
    [Arguments(101, "Cent un")]
    [Arguments(200, "Deux cents")]
    [Arguments(201, "Deux cent un")]
    [Arguments(500, "Cinq cents")]
    public async Task LFrFr003_CardinalText_Hundreds_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_fr_FR.GetListItemText("fr-FR", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText — thousands ─────────────────────────────────────────────

    [Test]
    [Arguments(1000, "Mille")]
    [Arguments(1001, "Mille un")]
    [Arguments(1100, "Mille cent")]
    [Arguments(2000, "Deux mille")]
    [Arguments(3000, "Trois mille")]
    public async Task LFrFr004_CardinalText_Thousands_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_fr_FR.GetListItemText("fr-FR", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinal ──────────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "1er")]
    [Arguments(2, "2e")]
    [Arguments(10, "10e")]
    [Arguments(21, "21e")]
    public async Task LFrFr005_Ordinal_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_fr_FR.GetListItemText("fr-FR", number, "ordinal");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText — special cases ──────────────────────────────────────────

    [Test]
    public async Task LFrFr006_OrdinalText_One_ReturnsPremier()
    {
        var result = ListItemTextGetter_fr_FR.GetListItemText("fr-FR", 1, "ordinalText");
        await Assert.That(result).IsEqualTo("Premier");
    }

    [Test]
    [Arguments(1000, "Millième")]
    public async Task LFrFr007_OrdinalText_Thousand_ReturnsMillieme(int number, string expected)
    {
        var result = ListItemTextGetter_fr_FR.GetListItemText("fr-FR", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText — basic values ───────────────────────────────────────────

    [Test]
    [Arguments(2, "Deuxième")]
    [Arguments(3, "Troisième")]
    [Arguments(10, "Dixième")]
    [Arguments(20, "Vingtième")]
    [Arguments(100, "Centième")]
    [Arguments(200, "Deux centième")]
    [Arguments(2000, "Deux millième")]
    public async Task LFrFr008_OrdinalText_TypicalValues_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_fr_FR.GetListItemText("fr-FR", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }
}
