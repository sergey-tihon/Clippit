// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word.Enums;

namespace Clippit.Word;

internal class ListItemTextGetter_Default
{
    private static readonly string[] OneThroughNineteen =
    [
        "one",
        "two",
        "three",
        "four",
        "five",
        "six",
        "seven",
        "eight",
        "nine",
        "ten",
        "eleven",
        "twelve",
        "thirteen",
        "fourteen",
        "fifteen",
        "sixteen",
        "seventeen",
        "eighteen",
        "nineteen",
    ];

    private static readonly string[] Tens =
    [
        "ten",
        "twenty",
        "thirty",
        "forty",
        "fifty",
        "sixty",
        "seventy",
        "eighty",
        "ninety",
    ];

    private static readonly string[] OrdinalOneThroughNineteen =
    [
        "first",
        "second",
        "third",
        "fourth",
        "fifth",
        "sixth",
        "seventh",
        "eighth",
        "ninth",
        "tenth",
        "eleventh",
        "twelfth",
        "thirteenth",
        "fourteenth",
        "fifteenth",
        "sixteenth",
        "seventeenth",
        "eighteenth",
        "nineteenth",
    ];

    private static readonly string[] OrdinalTenths =
    [
        "tenth",
        "twentieth",
        "thirtieth",
        "fortieth",
        "fiftieth",
        "sixtieth",
        "seventieth",
        "eightieth",
        "ninetieth",
    ];

    public static string GetListItemText(int levelNumber, NumberingFormatType numFmt)
    {
        switch (numFmt)
        {
            case NumberingFormatType.None:
                return "";
            case NumberingFormatType.Decimal:
                return levelNumber.ToString();
            case NumberingFormatType.DecimalZero when levelNumber <= 9:
                return "0" + levelNumber;
            case NumberingFormatType.DecimalZero:
                return levelNumber.ToString();
            case NumberingFormatType.UpperRoman:
                return RomanNumeralUtil.ToUpperRoman(levelNumber);
            case NumberingFormatType.LowerRoman:
                return RomanNumeralUtil.ToLowerRoman(levelNumber);
            case NumberingFormatType.UpperLetter:
            {
                var levelNumber2 = levelNumber % 780;
                if (levelNumber2 == 0)
                    levelNumber2 = 780;
                var a = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                var c = (levelNumber2 - 1) / 26;
                var n = (levelNumber2 - 1) % 26;
                var x = a[n];
                return "".PadRight(c + 1, x);
            }
            case NumberingFormatType.LowerLetter:
            {
                var levelNumber3 = levelNumber % 780;
                if (levelNumber3 == 0)
                    levelNumber3 = 780;
                var a = "abcdefghijklmnopqrstuvwxyz";
                var c = (levelNumber3 - 1) / 26;
                var n = (levelNumber3 - 1) % 26;
                var x = a[n];
                return "".PadRight(c + 1, x);
            }
            case NumberingFormatType.Ordinal:
            {
                string suffix;
                if (levelNumber % 100 == 11 || levelNumber % 100 == 12 || levelNumber % 100 == 13)
                    suffix = "th";
                else
                    suffix = (levelNumber % 10) switch
                    {
                        1 => "st",
                        2 => "nd",
                        3 => "rd",
                        _ => "th",
                    };
                return levelNumber + suffix;
            }
            case NumberingFormatType.CardinalText:
            {
                if (levelNumber <= 0 || levelNumber > 19999)
                    return levelNumber.ToString();
                var result = "";
                var t1 = levelNumber / 1000;
                var t2 = levelNumber % 1000;
                if (t1 >= 1)
                    result += OneThroughNineteen[t1 - 1] + " thousand";
                if (t1 >= 1 && t2 == 0)
                    return char.ToUpperInvariant(result[0]) + result[1..];
                if (t1 >= 1)
                    result += " ";
                var h1 = (levelNumber % 1000) / 100;
                var h2 = levelNumber % 100;
                if (h1 >= 1)
                    result += OneThroughNineteen[h1 - 1] + " hundred";
                if (h1 >= 1 && h2 == 0)
                    return char.ToUpperInvariant(result[0]) + result[1..];
                if (h1 >= 1)
                    result += " ";
                var z = levelNumber % 100;
                if (z <= 19)
                    result += OneThroughNineteen[z - 1];
                else
                {
                    var x = z / 10;
                    var r = z % 10;
                    result += Tens[x - 1];
                    if (r >= 1)
                        result += "-" + OneThroughNineteen[r - 1];
                }
                return char.ToUpperInvariant(result[0]) + result[1..];
            }
            case NumberingFormatType.OrdinalText:
            {
                if (levelNumber <= 0 || levelNumber > 19999)
                    return levelNumber.ToString();
                var result = "";
                var t1 = levelNumber / 1000;
                var t2 = levelNumber % 1000;
                if (t1 >= 1 && t2 != 0)
                    result += OneThroughNineteen[t1 - 1] + " thousand";
                if (t1 >= 1 && t2 == 0)
                {
                    result += OneThroughNineteen[t1 - 1] + " thousandth";
                    return char.ToUpperInvariant(result[0]) + result[1..];
                }
                if (t1 >= 1)
                    result += " ";
                var h1 = (levelNumber % 1000) / 100;
                var h2 = levelNumber % 100;
                if (h1 >= 1 && h2 != 0)
                    result += OneThroughNineteen[h1 - 1] + " hundred";
                if (h1 >= 1 && h2 == 0)
                {
                    result += OneThroughNineteen[h1 - 1] + " hundredth";
                    return char.ToUpperInvariant(result[0]) + result[1..];
                }
                if (h1 >= 1)
                    result += " ";
                var z = levelNumber % 100;
                if (z <= 19)
                    result += OrdinalOneThroughNineteen[z - 1];
                else
                {
                    var x = z / 10;
                    var r = z % 10;
                    if (r == 0)
                        result += OrdinalTenths[x - 1];
                    else
                        result += Tens[x - 1];
                    if (r >= 1)
                        result += "-" + OrdinalOneThroughNineteen[r - 1];
                }
                return char.ToUpperInvariant(result[0]) + result[1..];
            }
            case NumberingFormatType.DecimalPadded2:
                return $"{levelNumber:00}";
            case NumberingFormatType.DecimalPadded3:
                return $"{levelNumber:000}";
            case NumberingFormatType.DecimalPadded4:
                return $"{levelNumber:0000}";
            case NumberingFormatType.DecimalPadded5:
                return $"{levelNumber:00000}";
            case NumberingFormatType.Bullet:
                return "";
            case NumberingFormatType.DecimalEnclosedCircle when levelNumber >= 1 && levelNumber <= 20:
                return ((char)(9311 + levelNumber)).ToString();
            case NumberingFormatType.DecimalEnclosedCircle:
                return levelNumber.ToString();
            default:
                return levelNumber.ToString();
        }
    }
}
