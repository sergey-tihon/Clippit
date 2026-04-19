// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;
using Clippit.Word;

namespace Clippit.Tests.Word;

public class RomanNumeralUtilTests
{
    [Test]
    [Arguments(1, "I")]
    [Arguments(4, "IV")]
    [Arguments(9, "IX")]
    [Arguments(14, "XIV")]
    [Arguments(40, "XL")]
    [Arguments(90, "XC")]
    [Arguments(399, "CCCXCIX")]
    [Arguments(400, "CD")]
    [Arguments(900, "CM")]
    [Arguments(1000, "M")]
    [Arguments(1999, "MCMXCIX")]
    [Arguments(3999, "MMMCMXCIX")]
    [Arguments(10999, "MMMMMMMMMMCMXCIX")]
    public async Task RNU001_ToUpperRoman_ReturnsCorrectValue(int number, string expected)
    {
        var result = RomanNumeralUtil.ToUpperRoman(number);
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(1, "i")]
    [Arguments(4, "iv")]
    [Arguments(9, "ix")]
    [Arguments(40, "xl")]
    [Arguments(1999, "mcmxcix")]
    [Arguments(3999, "mmmcmxcix")]
    public async Task RNU002_ToLowerRoman_ReturnsLowercase(int number, string expected)
    {
        var result = RomanNumeralUtil.ToLowerRoman(number);
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(1, "i")]
    [Arguments(1999, "mcmxcix")]
    [Arguments(3999, "mmmcmxcix")]
    public async Task RNU003_ToLowerRoman_IsCultureInvariantUnderTurkishCulture(int number, string expected)
    {
        var original = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = new CultureInfo("tr-TR");
            var result = RomanNumeralUtil.ToLowerRoman(number);
            await Assert.That(result).IsEqualTo(expected);
        }
        finally
        {
            CultureInfo.CurrentCulture = original;
        }
    }

    [Test]
    [Arguments(0)]
    [Arguments(-1)]
    [Arguments(11000)]
    [Arguments(int.MinValue)]
    [Arguments(int.MaxValue)]
    public async Task RNU004_ToUpperRoman_ThrowsForOutOfRangeValues(int number)
    {
        await Assert.That(() => RomanNumeralUtil.ToUpperRoman(number)).Throws<ArgumentOutOfRangeException>();
    }

    [Test]
    [Arguments(0)]
    [Arguments(-1)]
    [Arguments(11000)]
    public async Task RNU005_ToLowerRoman_ThrowsForOutOfRangeValues(int number)
    {
        await Assert.That(() => RomanNumeralUtil.ToLowerRoman(number)).Throws<ArgumentOutOfRangeException>();
    }
}
