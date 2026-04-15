// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

public class ListItemTextRuRuTests
{
    // out-of-range guard — falls back to decimal string
    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    [Arguments(99999, "99999")]
    public async Task LRU000_OutOfRange_FallsBackToDecimal_CardinalText(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    public async Task LRU000b_OutOfRange_FallsBackToDecimal_OrdinalText(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0)]
    [Arguments(-1)]
    [Arguments(20000)]
    public async Task LRU000c_OutOfRange_UnsupportedNumFmt_ReturnsNull(int number)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "unsupportedFormat");
        await Assert.That(result).IsNull();
    }

    // cardinalText: 1–19 (OneThroughNineteen)
    [Test]
    [Arguments(1, "Один")]
    [Arguments(2, "Два")]
    [Arguments(9, "Девять")]
    [Arguments(10, "Десять")]
    [Arguments(11, "Одиннадцать")]
    [Arguments(19, "Девятнадцать")]
    public async Task LRU001_CardinalText_OneThroughNineteen(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // cardinalText: 20–99 (Tens ± units)
    [Test]
    [Arguments(20, "Двадцать")]
    [Arguments(21, "Двадцать один")]
    [Arguments(30, "Тридцать")]
    [Arguments(40, "Сорок")]
    [Arguments(50, "Пятьдесят")]
    [Arguments(99, "Девяносто девять")]
    public async Task LRU002_CardinalText_TwentyThroughNinetyNine(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // cardinalText: exact hundreds 100–900
    [Test]
    [Arguments(100, "Сто")]
    [Arguments(200, "Двести")]
    [Arguments(300, "Триста")]
    [Arguments(400, "Четыреста")]
    [Arguments(500, "Пятьсот")]
    [Arguments(600, "Шестьсот")]
    [Arguments(700, "Семьсот")]
    [Arguments(800, "Восемьсот")]
    [Arguments(900, "Девятьсот")]
    public async Task LRU003_CardinalText_ExactHundreds(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // cardinalText: compound hundreds
    [Test]
    [Arguments(101, "Сто один")]
    [Arguments(115, "Сто пятнадцать")]
    [Arguments(250, "Двести пятьдесят")]
    [Arguments(999, "Девятьсот девяносто девять")]
    public async Task LRU004_CardinalText_CompoundHundreds(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // cardinalText: exact thousands 1000–9000
    [Test]
    [Arguments(1000, "Одна тысяча")]
    [Arguments(2000, "Две тысячи")]
    [Arguments(3000, "Три тысячи")]
    [Arguments(4000, "Четыре тысячи")]
    [Arguments(5000, "Пять тысяч")]
    [Arguments(9000, "Девять тысяч")]
    public async Task LRU005_CardinalText_ExactThousands(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // cardinalText: compound thousands
    [Test]
    [Arguments(1001, "Одна тысяча один")]
    [Arguments(1100, "Одна тысяча сто")]
    [Arguments(2500, "Две тысячи пятьсот")]
    [Arguments(5050, "Пять тысяч пятьдесят")]
    public async Task LRU006_CardinalText_CompoundThousands(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ordinalText: 1–19
    [Test]
    [Arguments(1, "Первый")]
    [Arguments(2, "Второй")]
    [Arguments(3, "Третий")]
    [Arguments(5, "Пятый")]
    [Arguments(10, "Десятый")]
    [Arguments(19, "Девятнадцатый")]
    public async Task LRU007_OrdinalText_OneThroughNineteen(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ordinalText: exact tens 20–90
    [Test]
    [Arguments(20, "Двадцатый")]
    [Arguments(30, "Тридцатый")]
    [Arguments(40, "Сороковой")]
    [Arguments(50, "Пятидесятый")]
    [Arguments(90, "Девяностый")]
    public async Task LRU008_OrdinalText_ExactTens(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ordinalText: compound tens 21–99
    [Test]
    [Arguments(21, "Двадцать первый")]
    [Arguments(55, "Пятьдесят пятый")]
    [Arguments(99, "Девяносто девятый")]
    public async Task LRU009_OrdinalText_CompoundTens(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ordinalText: exact hundreds 100–900
    [Test]
    [Arguments(100, "Сотый")]
    [Arguments(200, "Двухсотый")]
    [Arguments(300, "Трёхсотый")]
    [Arguments(500, "Пятисотый")]
    [Arguments(900, "Девятисотый")]
    public async Task LRU010_OrdinalText_ExactHundreds(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ordinalText: compound hundreds
    [Test]
    [Arguments(101, "Сто первый")]
    [Arguments(250, "Двести пятидесятый")]
    [Arguments(999, "Девятьсот девяносто девятый")]
    public async Task LRU011_OrdinalText_CompoundHundreds(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ordinalText: exact thousands 1000–19000
    [Test]
    [Arguments(1000, "Тысячный")]
    [Arguments(2000, "Двухтысячный")]
    [Arguments(3000, "Трёхтысячный")]
    [Arguments(5000, "Пятитысячный")]
    [Arguments(9000, "Девятитысячный")]
    [Arguments(10000, "Десятитысячный")]
    [Arguments(11000, "Одиннадцатитысячный")]
    [Arguments(19000, "Девятнадцатитысячный")]
    public async Task LRU012_OrdinalText_ExactThousands(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ordinalText: compound thousands
    [Test]
    [Arguments(1001, "Одна тысяча первый")]
    [Arguments(1500, "Одна тысяча пятисотый")]
    [Arguments(2050, "Две тысячи пятидесятый")]
    public async Task LRU013_OrdinalText_CompoundThousands(int number, string expected)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // unsupported numFmt returns null
    [Test]
    [Arguments("decimal")]
    [Arguments("upperRoman")]
    [Arguments("ordinal")]
    public async Task LRU014_UnsupportedNumFmt_ReturnsNull(string numFmt)
    {
        var result = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", 1, numFmt);
        await Assert.That(result).IsNull();
    }

    // no English words appear in any output (regression guard)
    [Test]
    [Arguments(1)]
    [Arguments(19)]
    [Arguments(50)]
    [Arguments(100)]
    [Arguments(500)]
    [Arguments(999)]
    [Arguments(1000)]
    [Arguments(5000)]
    public async Task LRU015_CardinalAndOrdinal_ContainNoEnglishWords(int number)
    {
        string[] englishWords = ["thousand", "hundred", "hundredth", "thousandth"];

        var cardinal = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "cardinalText");
        var ordinal = ListItemTextGetter_ru_RU.GetListItemText("ru-RU", number, "ordinalText");

        foreach (var word in englishWords)
        {
            await Assert.That(cardinal).DoesNotContain(word);
            await Assert.That(ordinal).DoesNotContain(word);
        }
    }
}
