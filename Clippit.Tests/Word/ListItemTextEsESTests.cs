// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

public class ListItemTextEsESTests
{
    // ── Out-of-range guards ───────────────────────────────────────────────────

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    public async Task LES000_OutOfRange_FallsBackToDecimal_CardinalText(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    public async Task LES000b_OutOfRange_FallsBackToDecimal_OrdinalText(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0)]
    [Arguments(-1)]
    [Arguments(20000)]
    public async Task LES000c_UnsupportedFormat_ReturnsNull(int number)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "unsupportedFormat");
        await Assert.That(result).IsNull();
    }

    // ── cardinalText: 1–19 ───────────────────────────────────────────────────

    [Test]
    [Arguments(1, "Uno")]
    [Arguments(2, "Dos")]
    [Arguments(5, "Cinco")]
    [Arguments(10, "Diez")]
    [Arguments(11, "Once")]
    [Arguments(15, "Quince")]
    [Arguments(16, "Dieciséis")]
    [Arguments(19, "Diecinueve")]
    public async Task LES001_CardinalText_OneThroughNineteen(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText: 20–29 (fused veinti forms) ─────────────────────────────

    [Test]
    [Arguments(20, "Veinte")]
    [Arguments(21, "Veintiuno")]
    [Arguments(22, "Veintidós")]
    [Arguments(26, "Veintiséis")]
    [Arguments(29, "Veintinueve")]
    public async Task LES002_CardinalText_TwentyThroughTwentyNine(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText: 30–99 ──────────────────────────────────────────────────

    [Test]
    [Arguments(30, "Treinta")]
    [Arguments(31, "Treinta y uno")]
    [Arguments(40, "Cuarenta")]
    [Arguments(50, "Cincuenta")]
    [Arguments(99, "Noventa y nueve")]
    public async Task LES003_CardinalText_ThirtyThroughNinetyNine(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText: hundreds ───────────────────────────────────────────────

    [Test]
    [Arguments(100, "Cien")]
    [Arguments(101, "Ciento uno")]
    [Arguments(200, "Doscientos")]
    [Arguments(500, "Quinientos")]
    [Arguments(999, "Novecientos noventa y nueve")]
    public async Task LES004_CardinalText_Hundreds(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText: thousands ──────────────────────────────────────────────

    [Test]
    [Arguments(1000, "Mil")]
    [Arguments(2000, "Dos mil")]
    [Arguments(1500, "Mil quinientos")]
    public async Task LES005_CardinalText_Thousands(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText: 1–10 ────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "Primero")]
    [Arguments(2, "Segundo")]
    [Arguments(3, "Tercero")]
    [Arguments(7, "Séptimo")]
    [Arguments(10, "Décimo")]
    public async Task LES010_OrdinalText_OneThroughTen(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText: 11–19 ───────────────────────────────────────────────────

    [Test]
    [Arguments(11, "Undécimo")]
    [Arguments(12, "Duodécimo")]
    [Arguments(13, "Decimotercero")]
    [Arguments(19, "Decimonoveno")]
    public async Task LES011_OrdinalText_ElevenThroughNineteen(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText: 20–99 ───────────────────────────────────────────────────

    [Test]
    [Arguments(20, "Vigésimo")]
    [Arguments(21, "Vigésimo primero")]
    [Arguments(30, "Trigésimo")]
    [Arguments(32, "Trigésimo segundo")]
    [Arguments(40, "Cuadragésimo")]
    public async Task LES012_OrdinalText_TwentyThroughNinetyNine(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText: hundreds and thousands ─────────────────────────────────

    [Test]
    [Arguments(100, "Centésimo")]
    [Arguments(101, "Centésimo primero")]
    [Arguments(200, "Ducentésimo")]
    [Arguments(345, "Tricentésimo cuadragésimo quinto")]
    [Arguments(1000, "Milésimo")]
    [Arguments(2000, "Dos milésimo")]
    public async Task LES013_OrdinalText_HundredsAndThousands(int number, string expected)
    {
        var result = ListItemTextGetter_es_ES.GetListItemText("es-ES", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }
}
