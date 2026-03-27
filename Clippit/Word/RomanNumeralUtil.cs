// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Word;

internal static class RomanNumeralUtil
{
    private static readonly string[] s_ones = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX"];
    private static readonly string[] s_tens = ["", "X", "XX", "XXX", "XL", "L", "LX", "LXX", "LXXX", "XC"];
    private static readonly string[] s_hundreds = ["", "C", "CC", "CCC", "CD", "D", "DC", "DCC", "DCCC", "CM", "M"];
    private static readonly string[] s_thousands =
    [
        "",
        "M",
        "MM",
        "MMM",
        "MMMM",
        "MMMMM",
        "MMMMMM",
        "MMMMMMM",
        "MMMMMMMM",
        "MMMMMMMMM",
        "MMMMMMMMMM",
    ];

    public static string ToUpperRoman(int number)
    {
        var ones = number % 10;
        var tens = (number % 100) / 10;
        var hundreds = (number % 1000) / 100;
        var thousands = number / 1000;
        return s_thousands[thousands] + s_hundreds[hundreds] + s_tens[tens] + s_ones[ones];
    }

    public static string ToLowerRoman(int number) => ToUpperRoman(number).ToLowerInvariant();
}
