// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Globalization;

namespace Clippit.Word
{
    public class ListItemTextGetter_tr_TR
    {
        private static readonly string[] OneThroughNineteen =
        {
            "bir",
            "iki",
            "üç",
            "dört",
            "beş",
            "altı",
            "yedi",
            "sekiz",
            "dokuz",
            "on",
            "onbir",
            "oniki",
            "onüç",
            "ondört",
            "onbeş",
            "onaltı",
            "onyedi",
            "onsekiz",
            "ondokuz",
        };

        private static readonly string[] Tens =
        {
            "on",
            "yirmi",
            "otuz",
            "kırk",
            "elli",
            "altmış",
            "yetmiş",
            "seksen",
            "doksan",
        };

        private static readonly string[] OrdinalOneThroughNineteen =
        {
            "birinci",
            "ikinci",
            "üçüncü",
            "dördüncü",
            "beşinci",
            "altıncı",
            "yedinci",
            "sekizinci",
            "dokuzuncu",
            "onuncu",
            "onbirinci",
            "onikinci",
            "onüçüncü",
            "ondördüncü",
            "onbeşinci",
            "onaltıncı",
            "onyedinci",
            "onsekizinci",
            "ondokuzuncu",
        };

        private static readonly string[] TwoThroughNineteen =
        {
            "",
            "iki",
            "üç",
            "dört",
            "beş",
            "altı",
            "yedi",
            "sekiz",
            "dokuz",
            "on",
            "onbir",
            "oniki",
            "onüç",
            "ondört",
            "onbeş",
            "onaltı",
            "onyedi",
            "onsekiz",
            "ondokuz",
        };

        private static readonly string[] OrdinalTenths =
        {
            "onuncu",
            "yirminci",
            "otuzuncu",
            "kırkıncı",
            "ellinci",
            "altmışıncı",
            "yetmişinci",
            "sekseninci",
            "doksanıncı",
        };

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
            #region
            if (numFmt == "decimal")
            {
                return levelNumber.ToString();
            }
            if (numFmt == "decimalZero")
            {
                if (levelNumber <= 9)
                    return "0" + levelNumber;
                else
                    return levelNumber.ToString();
            }
            if (numFmt == "upperRoman")
            {
                return RomanNumeralUtil.ToUpperRoman(levelNumber);
            }
            if (numFmt == "lowerRoman")
            {
                return RomanNumeralUtil.ToLowerRoman(levelNumber);
            }
            if (numFmt == "upperLetter")
            {
                var a = "ABCÇDEFGĞHIİJKLMNOÖPRSŞTUÜVYZ";
                //string a = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                var c = (levelNumber - 1) / 29;
                var n = (levelNumber - 1) % 29;
                var x = a[n];
                return "".PadRight(c + 1, x);
            }
            if (numFmt == "lowerLetter")
            {
                var a = "abcçdefgğhıijklmnoöprsştuüvyz";
                var c = (levelNumber - 1) / 29;
                var n = (levelNumber - 1) % 29;
                var x = a[n];
                return "".PadRight(c + 1, x);
            }
            if (numFmt == "ordinal")
            {
                var suffix =
                    /*if (levelNumber % 100 == 11 || levelNumber % 100 == 12 ||
                    levelNumber % 100 == 13)
                    suffix = "th";
                else if (levelNumber % 10 == 1)
                    suffix = "st";
                else if (levelNumber % 10 == 2)
                    suffix = "nd";
                else if (levelNumber % 10 == 3)
                    suffix = "rd";
                else
                    suffix = "th";*/
                    ".";
                return levelNumber + suffix;
            }
            if (numFmt == "cardinalText")
            {
                if (levelNumber <= 0 || levelNumber > 19999)
                    return levelNumber.ToString();
                var result = "";
                var t1 = levelNumber / 1000;
                var t2 = levelNumber % 1000;
                if (t1 >= 1)
                    result += OneThroughNineteen[t1 - 1] + " bin";
                if (t1 >= 1 && t2 == 0)
                    return result.Substring(0, 1).ToUpper(CultureInfo.GetCultureInfo("tr-TR")) + result.Substring(1);
                if (t1 >= 1)
                    result += " ";
                var h1 = (levelNumber % 1000) / 100;
                var h2 = levelNumber % 100;
                if (h1 >= 1)
                    result += OneThroughNineteen[h1 - 1] + " yüz";
                if (h1 >= 1 && h2 == 0)
                    return result.Substring(0, 1).ToUpper(CultureInfo.GetCultureInfo("tr-TR")) + result.Substring(1);
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
                        result += /*"-" + */
                        OneThroughNineteen[r - 1];
                }
                return result.Substring(0, 1).ToUpper(CultureInfo.GetCultureInfo("tr-TR")) + result.Substring(1);
            }
            #endregion
            if (numFmt == "ordinalText")
            {
                if (levelNumber <= 0 || levelNumber > 19999)
                    return levelNumber.ToString();
                var result = "";
                var t1 = levelNumber / 1000;
                var t2 = levelNumber % 1000;
                if (t1 >= 1 && t2 != 0)
                    result += TwoThroughNineteen[t1 - 1] + "bin";
                if (t1 >= 1 && t2 == 0)
                {
                    result += TwoThroughNineteen[t1 - 1] + "bininci";
                    return result.Substring(0, 1).ToUpper(CultureInfo.GetCultureInfo("tr-TR")) + result.Substring(1);
                }
                //if (t1 >= 1)
                //    result += " ";
                var h1 = (levelNumber % 1000) / 100;
                var h2 = levelNumber % 100;
                if (h1 >= 1 && h2 != 0)
                    result += TwoThroughNineteen[h1 - 1] + "yüz";
                if (h1 >= 1 && h2 == 0)
                {
                    result += TwoThroughNineteen[h1 - 1] + "yüzüncü";
                    return result.Substring(0, 1).ToUpper(CultureInfo.GetCultureInfo("tr-TR")) + result.Substring(1);
                }
                //if (h1 >= 1)
                //    result += " ";
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
                        result += OrdinalOneThroughNineteen[r - 1]; //result += "-" + OrdinalOneThroughNineteen[r - 1];
                }
                return result.Substring(0, 1).ToUpper(CultureInfo.GetCultureInfo("tr-TR")) + result.Substring(1);
            }
            if (numFmt == "0001, 0002, 0003, ...")
            {
                return $"{levelNumber:0000}";
            }
            if (numFmt == "bullet")
                return "";
            return levelNumber.ToString();
        }
    }
}
