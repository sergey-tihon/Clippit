// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

/// <summary>
/// Unit tests for <see cref="ListItemTextGetter_Default.GetListItemText"/>.
/// This is the primary list-numbering implementation for all non-locale-specific formats.
/// The class is internal and exposed via InternalsVisibleTo("Clippit.Tests").
/// </summary>
public class ListItemTextDefaultTests
{
    // ── none ────────────────────────────────────────────────────────────────

    [Test]
    public async Task LDef001_None_ReturnsEmpty()
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", 5, "none");
        await Assert.That(result).IsEqualTo("");
    }

    // ── decimal ─────────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "1")]
    [Arguments(9, "9")]
    [Arguments(10, "10")]
    [Arguments(100, "100")]
    public async Task LDef002_Decimal_ReturnsNumberAsString(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "decimal");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── decimalZero ──────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "01")]
    [Arguments(9, "09")]
    [Arguments(10, "10")]
    [Arguments(99, "99")]
    [Arguments(100, "100")]
    public async Task LDef003_DecimalZero_PadsSingleDigitsWithLeadingZero(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "decimalZero");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── upperRoman ───────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "I")]
    [Arguments(4, "IV")]
    [Arguments(9, "IX")]
    [Arguments(40, "XL")]
    [Arguments(1999, "MCMXCIX")]
    public async Task LDef004_UpperRoman_ReturnsUppercaseRomanNumerals(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "upperRoman");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── lowerRoman ───────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "i")]
    [Arguments(4, "iv")]
    [Arguments(9, "ix")]
    [Arguments(40, "xl")]
    [Arguments(1999, "mcmxcix")]
    public async Task LDef005_LowerRoman_ReturnsLowercaseRomanNumerals(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "lowerRoman");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── upperLetter ──────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "A")]
    [Arguments(26, "Z")]
    [Arguments(27, "AA")]
    [Arguments(52, "ZZ")]
    [Arguments(53, "AAA")]
    public async Task LDef006_UpperLetter_SequenceAtoZThenDoubled(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "upperLetter");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    public async Task LDef006b_UpperLetter_WrapAroundAt780_Restarts()
    {
        // After 780 the sequence wraps: 781 → "A" again
        var result = ListItemTextGetter_Default.GetListItemText("en-US", 781, "upperLetter");
        await Assert.That(result).IsEqualTo("A");
    }

    // ── lowerLetter ──────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "a")]
    [Arguments(26, "z")]
    [Arguments(27, "aa")]
    [Arguments(52, "zz")]
    [Arguments(53, "aaa")]
    public async Task LDef007_LowerLetter_SequenceAtoZThenDoubled(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "lowerLetter");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinal ──────────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "1st")]
    [Arguments(2, "2nd")]
    [Arguments(3, "3rd")]
    [Arguments(4, "4th")]
    [Arguments(11, "11th")] // teen exception: 11th not 11st
    [Arguments(12, "12th")] // teen exception: 12th not 12nd
    [Arguments(13, "13th")] // teen exception: 13th not 13rd
    [Arguments(21, "21st")]
    [Arguments(22, "22nd")]
    [Arguments(23, "23rd")]
    [Arguments(100, "100th")]
    [Arguments(111, "111th")] // 111 % 100 = 11 → teen exception
    [Arguments(121, "121st")]
    public async Task LDef008_Ordinal_AppliesCorrectSuffix(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "ordinal");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── cardinalText ─────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "One")]
    [Arguments(2, "Two")]
    [Arguments(10, "Ten")]
    [Arguments(11, "Eleven")]
    [Arguments(19, "Nineteen")]
    [Arguments(20, "Twenty")]
    [Arguments(21, "Twenty-one")]
    [Arguments(99, "Ninety-nine")]
    [Arguments(100, "One hundred")]
    [Arguments(200, "Two hundred")]
    [Arguments(110, "One hundred ten")]
    [Arguments(121, "One hundred twenty-one")]
    [Arguments(1000, "One thousand")]
    [Arguments(1001, "One thousand one")]
    [Arguments(1100, "One thousand one hundred")]
    [Arguments(2019, "Two thousand nineteen")]
    public async Task LDef009_CardinalText_EnglishWords(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    [Arguments(99999, "99999")]
    public async Task LDef009b_CardinalText_OutOfRange_FallsBackToDecimal(int number, string expected)
    {
        // levelNumber == 0 and levelNumber >= 20000 previously caused IndexOutOfRangeException
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "cardinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ordinalText ──────────────────────────────────────────────────────────

    [Test]
    [Arguments(1, "First")]
    [Arguments(2, "Second")]
    [Arguments(3, "Third")]
    [Arguments(10, "Tenth")]
    [Arguments(11, "Eleventh")]
    [Arguments(19, "Nineteenth")]
    [Arguments(20, "Twentieth")]
    [Arguments(21, "Twenty-first")]
    [Arguments(30, "Thirtieth")]
    [Arguments(100, "One hundredth")]
    [Arguments(200, "Two hundredth")]
    [Arguments(1000, "One thousandth")]
    [Arguments(2000, "Two thousandth")]
    public async Task LDef010_OrdinalText_EnglishOrdinalWords(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(0, "0")]
    [Arguments(-1, "-1")]
    [Arguments(20000, "20000")]
    [Arguments(99999, "99999")]
    public async Task LDef010b_OrdinalText_OutOfRange_FallsBackToDecimal(int number, string expected)
    {
        // levelNumber == 0 and levelNumber >= 20000 previously caused IndexOutOfRangeException
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "ordinalText");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── zero-padded custom formats ───────────────────────────────────────────

    [Test]
    [Arguments(5, "01, 02, 03, ...", "05")]
    [Arguments(12, "01, 02, 03, ...", "12")]
    [Arguments(5, "001, 002, 003, ...", "005")]
    [Arguments(5, "0001, 0002, 0003, ...", "0005")]
    [Arguments(5, "00001, 00002, 00003, ...", "00005")]
    public async Task LDef011_ZeroPaddedFormats_PadToCorrectWidth(int number, string numFmt, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, numFmt);
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── bullet ───────────────────────────────────────────────────────────────
    // For "bullet" numFmt the actual bullet character is defined by <w:lvlText>,
    // not by the counter; GetListItemText therefore returns "" for this format.

    [Test]
    public async Task LDef012_Bullet_ReturnsEmpty()
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", 1, "bullet");
        await Assert.That(result).IsEqualTo("");
    }

    // ── decimalEnclosedCircle ────────────────────────────────────────────────

    [Test]
    [Arguments(1, "\u2460")] // ① (Unicode CIRCLED DIGIT ONE)
    [Arguments(10, "\u2469")] // ⑩
    [Arguments(20, "\u2473")] // ⑳
    public async Task LDef013_DecimalEnclosedCircle_ReturnsCircledDigitForRange1To20(int number, string expected)
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", number, "decimalEnclosedCircle");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    public async Task LDef013b_DecimalEnclosedCircle_OutOfRange_FallsBackToDecimal()
    {
        // Numbers outside 1–20 fall back to plain decimal
        var result = ListItemTextGetter_Default.GetListItemText("en-US", 21, "decimalEnclosedCircle");
        await Assert.That(result).IsEqualTo("21");
    }

    // ── unknown / default ────────────────────────────────────────────────────

    [Test]
    public async Task LDef014_UnknownFormat_FallsBackToDecimal()
    {
        var result = ListItemTextGetter_Default.GetListItemText("en-US", 42, "unknownFormat");
        await Assert.That(result).IsEqualTo("42");
    }
}
