// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

/// <summary>
/// Unit tests for <see cref="ListItemTextGetter_sv_SE.GetListItemText"/>.
/// </summary>
public class ListItemTextSvSeTests
{
    // ── cardinalText — boundary conditions ───────────────────────────────────

    [Test]
    public async Task LSvSe001_CardinalText_Zero_ReturnsNoll()
    {
        var result = ListItemTextGetter_sv_SE.GetListItemText("sv-SE", 0, "cardinalText");
        await Assert.That(result).IsEqualTo("Noll");
    }

    [Test]
    public async Task LSvSe002_CardinalText_NegativeNumber_Throws()
    {
        await Assert.That(() => ListItemTextGetter_sv_SE.GetListItemText("sv-SE", -1, "cardinalText"))
            .Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task LSvSe003_CardinalText_Negative_Throws()
    {
        await Assert.That(() => ListItemTextGetter_sv_SE.GetListItemText("sv-SE", -5, "cardinalText"))
            .Throws<ArgumentOutOfRangeException>();
    }

    // ── cardinalText — ones / teens ───────────────────────────────────────────

    [Test]
    [Arguments(1, "Ett")]
    [Arguments(2, "Två")]
    [Arguments(9, "Nio")]
    [Arguments(10, "Tio")]
    [Arguments(11, "Elva")]
    [Arguments(19, "Nitton")]
    public async Task LSvSe004_CardinalText_OnesToNineteen_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_sv_SE.GetListItemText("sv-SE", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText — tens ───────────────────────────────────────────────────

    [Test]
    [Arguments(20, "Tjugo")]
    [Arguments(21, "Tjugoett")]
    [Arguments(30, "Trettio")]
    [Arguments(99, "Nittionio")]
    public async Task LSvSe005_CardinalText_Tens_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_sv_SE.GetListItemText("sv-SE", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText — hundreds / thousands ──────────────────────────────────

    [Test]
    [Arguments(100, "Etthundra")]
    [Arguments(101, "Etthundraett")]
    [Arguments(200, "Tvåhundra")]
    [Arguments(1000, "Ettusen")]
    [Arguments(1001, "Ettusenett")]
    [Arguments(2000, "Tvåtusen")]
    public async Task LSvSe006_CardinalText_HundredsAndThousands_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_sv_SE.GetListItemText("sv-SE", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText — boundary conditions ────────────────────────────────────

    [Test]
    public async Task LSvSe007_OrdinalText_Zero_Throws()
    {
        await Assert.That(() => ListItemTextGetter_sv_SE.GetListItemText("sv-SE", 0, "ordinalText"))
            .Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    public async Task LSvSe008_OrdinalText_TooLarge_Throws()
    {
        await Assert.That(() => ListItemTextGetter_sv_SE.GetListItemText("sv-SE", 10000, "ordinalText"))
            .Throws<ArgumentOutOfRangeException>();
    }

    // ── ordinalText — special case ────────────────────────────────────────────

    [Test]
    public async Task LSvSe009_OrdinalText_One_ReturnsFörsta()
    {
        var result = ListItemTextGetter_sv_SE.GetListItemText("sv-SE", 1, "ordinalText");
        await Assert.That(result).IsEqualTo("Första");
    }

    // ── ordinalText — basic values ────────────────────────────────────────────

    [Test]
    [Arguments(2, "Andra")]
    [Arguments(3, "Tredje")]
    [Arguments(10, "Tionde")]
    [Arguments(20, "Tjugonde")]
    [Arguments(100, "Etthundrade")]
    [Arguments(1000, "Ettusende")]
    [Arguments(2000, "Tvåtusende")]
    public async Task LSvSe010_OrdinalText_TypicalValues_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_sv_SE.GetListItemText("sv-SE", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinal ───────────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "1:a")]
    [Arguments(2, "2:a")]
    [Arguments(3, "3:e")]
    [Arguments(11, "11:a")]
    [Arguments(21, "21:a")]
    [Arguments(22, "22:a")]
    [Arguments(23, "23:e")]
    public async Task LSvSe011_Ordinal_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_sv_SE.GetListItemText("sv-SE", number, "ordinal");
        await Assert.That(result).IsEqualTo(expected);
    }
}
