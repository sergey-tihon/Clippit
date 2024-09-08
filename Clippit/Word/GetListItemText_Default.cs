// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Word
{
    internal class ListItemTextGetter_Default
    {
        private static readonly string[] RomanOnes =
        {
            "",
            "I",
            "II",
            "III",
            "IV",
            "V",
            "VI",
            "VII",
            "VIII",
            "IX",
        };

        private static readonly string[] RomanTens =
        {
            "",
            "X",
            "XX",
            "XXX",
            "XL",
            "L",
            "LX",
            "LXX",
            "LXXX",
            "XC",
        };

        private static readonly string[] RomanHundreds =
        {
            "",
            "C",
            "CC",
            "CCC",
            "CD",
            "D",
            "DC",
            "DCC",
            "DCCC",
            "CM",
            "M",
        };

        private static readonly string[] RomanThousands =
        {
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
        };

        private static readonly string[] OneThroughNineteen =
        {
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
        };

        private static readonly string[] Tens =
        {
            "ten",
            "twenty",
            "thirty",
            "forty",
            "fifty",
            "sixty",
            "seventy",
            "eighty",
            "ninety",
        };

        private static readonly string[] OrdinalOneThroughNineteen =
        {
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
        };

        private static readonly string[] OrdinalTenths =
        {
            "tenth",
            "twentieth",
            "thirtieth",
            "fortieth",
            "fiftieth",
            "sixtieth",
            "seventieth",
            "eightieth",
            "ninetieth",
        };

        public static string GetListItemText(
            string languageCultureName,
            int levelNumber,
            string numFmt
        )
        {
            switch (numFmt)
            {
                case "none":
                    return "";
                case "decimal":
                    return levelNumber.ToString();
                case "decimalZero" when levelNumber <= 9:
                    return "0" + levelNumber;
                case "decimalZero":
                    return levelNumber.ToString();
                case "upperRoman":
                {
                    var ones = levelNumber % 10;
                    var tens = (levelNumber % 100) / 10;
                    var hundreds = (levelNumber % 1000) / 100;
                    var thousands = levelNumber / 1000;
                    return RomanThousands[thousands]
                        + RomanHundreds[hundreds]
                        + RomanTens[tens]
                        + RomanOnes[ones];
                }
                case "lowerRoman":
                {
                    var ones = levelNumber % 10;
                    var tens = (levelNumber % 100) / 10;
                    var hundreds = (levelNumber % 1000) / 100;
                    var thousands = levelNumber / 1000;
                    return (
                        RomanThousands[thousands]
                        + RomanHundreds[hundreds]
                        + RomanTens[tens]
                        + RomanOnes[ones]
                    ).ToLower();
                }
                case "upperLetter":
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
                case "lowerLetter":
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
                case "ordinal":
                {
                    string suffix;
                    if (
                        levelNumber % 100 == 11
                        || levelNumber % 100 == 12
                        || levelNumber % 100 == 13
                    )
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
                case "cardinalText":
                {
                    var result = "";
                    var t1 = levelNumber / 1000;
                    var t2 = levelNumber % 1000;
                    if (t1 >= 1)
                        result += OneThroughNineteen[t1 - 1] + " thousand";
                    if (t1 >= 1 && t2 == 0)
                        return result.Substring(0, 1).ToUpper() + result.Substring(1);
                    if (t1 >= 1)
                        result += " ";
                    var h1 = (levelNumber % 1000) / 100;
                    var h2 = levelNumber % 100;
                    if (h1 >= 1)
                        result += OneThroughNineteen[h1 - 1] + " hundred";
                    if (h1 >= 1 && h2 == 0)
                        return result.Substring(0, 1).ToUpper() + result.Substring(1);
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
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }
                case "ordinalText":
                {
                    var result = "";
                    var t1 = levelNumber / 1000;
                    var t2 = levelNumber % 1000;
                    if (t1 >= 1 && t2 != 0)
                        result += OneThroughNineteen[t1 - 1] + " thousand";
                    if (t1 >= 1 && t2 == 0)
                    {
                        result += OneThroughNineteen[t1 - 1] + " thousandth";
                        return result.Substring(0, 1).ToUpper() + result.Substring(1);
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
                        return result.Substring(0, 1).ToUpper() + result.Substring(1);
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
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }
                case "01, 02, 03, ...":
                    return $"{levelNumber:00}";
                case "001, 002, 003, ...":
                    return $"{levelNumber:000}";
                case "0001, 0002, 0003, ...":
                    return $"{levelNumber:0000}";
                case "00001, 00002, 00003, ...":
                    return $"{levelNumber:00000}";
                case "bullet":
                    return "";
                case "decimalEnclosedCircle" when levelNumber >= 1 && levelNumber <= 20:
                {
                    // 9311 + levelNumber
                    var s = new string(new[] { (char)(9311 + levelNumber) });
                    return s;
                }
                case "decimalEnclosedCircle":
                    return levelNumber.ToString();
                default:
                    return levelNumber.ToString();
            }
        }
    }
}
