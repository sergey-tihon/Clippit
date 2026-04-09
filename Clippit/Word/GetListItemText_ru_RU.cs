// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Word
{
    public class ListItemTextGetter_ru_RU
    {
        private static readonly string[] OneThroughNineteen =
        {
            "один",
            "два",
            "три",
            "четыре",
            "пять",
            "шесть",
            "семь",
            "восемь",
            "девять",
            "десять",
            "одиннадцать",
            "двенадцать",
            "тринадцать",
            "четырнадцать",
            "пятнадцать",
            "шестнадцать",
            "семнадцать",
            "восемнадцать",
            "девятнадцать",
        };

        // Tens for 20–90 (index 0 = двадцать, index 7 = девяносто).
        // Used with x = z/10 where z ∈ [20,99]: Tens[x - 2].
        private static readonly string[] Tens =
        {
            "двадцать",
            "тридцать",
            "сорок",
            "пятьдесят",
            "шестьдесят",
            "семьдесят",
            "восемьдесят",
            "девяносто",
        };

        // Hundreds for 100–900 (index 0 = сто, index 8 = девятьсот).
        // Used with h1 ∈ [1,9]: Hundreds[h1 - 1].
        private static readonly string[] Hundreds =
        {
            "сто",
            "двести",
            "триста",
            "четыреста",
            "пятьсот",
            "шестьсот",
            "семьсот",
            "восемьсот",
            "девятьсот",
        };

        private static readonly string[] OrdinalOneThroughNineteen =
        {
            "первый",
            "второй",
            "третий",
            "четвертый",
            "пятый",
            "шестой",
            "седьмой",
            "восьмой",
            "девятый",
            "десятый",
            "одиннадцатый",
            "двенадцатый",
            "тринадцатый",
            "четырнадцатый",
            "пятнадцатый",
            "шестнадцатый",
            "семнадцатый",
            "восемнадцатый",
            "девятнадцатый",
        };

        // Ordinal tens for 20th–90th (index 0 = двадцатый, index 7 = девяностый).
        // Used with x = z/10 where z ∈ [20,99] and z%10 == 0: OrdinalTens[x - 2].
        private static readonly string[] OrdinalTens =
        {
            "двадцатый",
            "тридцатый",
            "сороковой",
            "пятидесятый",
            "шестидесятый",
            "семидесятый",
            "восьмидесятый",
            "девяностый",
        };

        // Ordinal hundreds for 100th–900th (index 0 = сотый, index 8 = девятисотый).
        // Used with h1 ∈ [1,9]: OrdinalHundreds[h1 - 1].
        private static readonly string[] OrdinalHundreds =
        {
            "сотый",
            "двухсотый",
            "трёхсотый",
            "четырёхсотый",
            "пятисотый",
            "шестисотый",
            "семисотый",
            "восьмисотый",
            "девятисотый",
        };

        private static string CardinalThousands(int t1) =>
            t1 switch
            {
                1 => "одна тысяча",
                2 => "две тысячи",
                3 or 4 => OneThroughNineteen[t1 - 1] + " тысячи",
                _ => OneThroughNineteen[t1 - 1] + " тысяч",
            };

        private static string OrdinalThousands(int t1) =>
            t1 switch
            {
                1 => "тысячный",
                2 => "двухтысячный",
                3 => "трёхтысячный",
                4 => "четырёхтысячный",
                5 => "пятитысячный",
                6 => "шеститысячный",
                7 => "семитысячный",
                8 => "восьмитысячный",
                9 => "девятитысячный",
                10 => "десятитысячный",
                11 => "одиннадцатитысячный",
                12 => "двенадцатитысячный",
                13 => "тринадцатитысячный",
                14 => "четырнадцатитысячный",
                15 => "пятнадцатитысячный",
                16 => "шестнадцатитысячный",
                17 => "семнадцатитысячный",
                18 => "восемнадцатитысячный",
                19 => "девятнадцатитысячный",
                _ => OneThroughNineteen[t1 - 1] + "тысячный",
            };

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
            if (numFmt == "cardinalText")
            {
                var result = "";
                var t1 = levelNumber / 1000;
                var t2 = levelNumber % 1000;
                if (t1 >= 1)
                    result += CardinalThousands(t1);
                if (t1 >= 1 && t2 == 0)
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                if (t1 >= 1)
                    result += " ";
                var h1 = (levelNumber % 1000) / 100;
                var h2 = levelNumber % 100;
                if (h1 >= 1)
                    result += Hundreds[h1 - 1];
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
                    result += Tens[x - 2];
                    if (r >= 1)
                        result += " " + OneThroughNineteen[r - 1];
                }
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            if (numFmt == "ordinalText")
            {
                var result = "";
                var t1 = levelNumber / 1000;
                var t2 = levelNumber % 1000;
                if (t1 >= 1 && t2 != 0)
                    result += CardinalThousands(t1);
                if (t1 >= 1 && t2 == 0)
                {
                    result += OrdinalThousands(t1);
                    return result.Substring(0, 1).ToUpper() + result.Substring(1);
                }
                if (t1 >= 1)
                    result += " ";
                var h1 = (levelNumber % 1000) / 100;
                var h2 = levelNumber % 100;
                if (h1 >= 1 && h2 != 0)
                    result += Hundreds[h1 - 1];
                if (h1 >= 1 && h2 == 0)
                {
                    result += OrdinalHundreds[h1 - 1];
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
                        result += OrdinalTens[x - 2];
                    else
                        result += Tens[x - 2];
                    if (r >= 1)
                        result += " " + OrdinalOneThroughNineteen[r - 1];
                }
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            return null;
        }
    }
}
