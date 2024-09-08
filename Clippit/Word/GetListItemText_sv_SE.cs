﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;

namespace Clippit.Word
{
    public class ListItemTextGetter_sv_SE
    {
        private static readonly string[] OneThroughNineteen =
        {
            "",
            "ett",
            "två",
            "tre",
            "fyra",
            "fem",
            "sex",
            "sju",
            "åtta",
            "nio",
            "tio",
            "elva",
            "tolv",
            "tretton",
            "fjorton",
            "femton",
            "sexton",
            "sjutton",
            "arton",
            "nitton",
        };

        private static readonly string[] Tens =
        {
            "",
            "tio",
            "tjugo",
            "trettio",
            "fyrtio",
            "femtio",
            "sextio",
            "sjuttio",
            "åttio",
            "nittio",
            "etthundra",
        };

        private static readonly string[] OrdinalOneThroughNineteen =
        {
            "",
            "första",
            "andra",
            "tredje",
            "fjärde",
            "femte",
            "sjätte",
            "sjunde",
            "åttonde",
            "nionde",
            "tionde",
            "elfte",
            "tolfte",
            "trettonde",
            "fjortonde",
            "femtonde",
            "sextonde",
            "sjuttonde",
            "artonde",
            "nittonde",
        };

        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
            return numFmt switch
            {
                "cardinalText" => NumberAsCardinalText(languageCultureName, levelNumber, numFmt),
                "ordinalText" => NumberAsOrdinalText(languageCultureName, levelNumber, numFmt),
                "ordinal" => NumberAsOrdinal(languageCultureName, levelNumber, numFmt),
                _ => null,
            };
        }

        private static string NumberAsCardinalText(string languageCultureName, int levelNumber, string numFmt)
        {
            var result = "";

            var sLevel = (levelNumber + 10000).ToString();
            var thousands = int.Parse(sLevel.Substring(1, 1));
            var hundreds = int.Parse(sLevel.Substring(2, 1));
            var tens = int.Parse(sLevel.Substring(3, 1));
            var ones = int.Parse(sLevel.Substring(4, 1));

            //Validation
            if (thousands > 19)
                throw new ArgumentOutOfRangeException(
                    nameof(levelNumber),
                    "Convering a levelNumber to ordinal text that is greater then 19 999 is not supported"
                );
            if (levelNumber == 0)
                return "Noll";
            if (levelNumber < 0)
                throw new ArgumentOutOfRangeException(
                    nameof(levelNumber),
                    "Converting a negative levelNumber to ordinal text is not supported"
                );

            /* exact thousands */
            if (levelNumber == 1000)
                return "Ettusen";
            if (levelNumber > 1000 && hundreds == 0 && tens == 0 && ones == 0)
            {
                result = OneThroughNineteen[thousands] + "tusen";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }

            /* > 1000 */
            if (levelNumber > 1000 && levelNumber < 2000)
                result = "ettusen";
            else if (levelNumber > 2000 && levelNumber < 10000)
                result = OneThroughNineteen[thousands] + "tusen";

            /* exact hundreds */
            if (hundreds > 0 && tens == 0 && ones == 0)
            {
                if (hundreds == 1)
                    result += "etthundra";
                else
                    result += OneThroughNineteen[hundreds] + "hundra";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }

            /* > 100 */
            if (hundreds > 0)
            {
                if (hundreds == 1)
                    result += "etthundra";
                else
                    result += OneThroughNineteen[hundreds] + "hundra";
            }

            /* exact tens */
            if (tens > 0 && ones == 0)
            {
                result += Tens[tens];
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }

            /* > 20 */
            if (tens == 1)
            {
                result += OneThroughNineteen[tens * 10 + ones];
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            else if (tens > 1)
            {
                result += Tens[tens] + OneThroughNineteen[ones];
                ;
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            else
            {
                result += OneThroughNineteen[ones];
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
        }

        private static string NumberAsOrdinalText(string languageCultureName, int levelNumber, string numFmt)
        {
            var result = "";

            if (levelNumber <= 0)
                throw new ArgumentOutOfRangeException(
                    nameof(levelNumber),
                    "Converting a zero or negative levelNumber to ordinal text is not supported"
                );
            if (levelNumber >= 10000)
                throw new ArgumentOutOfRangeException(
                    nameof(levelNumber),
                    "Convering a levelNumber to ordinal text that is greater then 10000 is not supported"
                );

            if (levelNumber == 1)
                return "Första";

            var sLevel = (levelNumber + 10000).ToString();
            var thousands = int.Parse(sLevel.Substring(1, 1));
            var hundreds = int.Parse(sLevel.Substring(2, 1));
            var tens = int.Parse(sLevel.Substring(3, 1));
            var ones = int.Parse(sLevel.Substring(4, 1));

            /* exact thousands */
            if (levelNumber == 1000)
                return "Ettusende";
            if (levelNumber > 1000 && hundreds == 0 && tens == 0 && ones == 0)
            {
                result = OneThroughNineteen[thousands] + "tusende";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }

            /* > 1000 */
            if (levelNumber > 1000 && levelNumber < 2000)
                result = "ettusen";
            else if (levelNumber > 2000 && levelNumber < 10000)
                result = OneThroughNineteen[thousands] + "tusende";

            /* exact hundreds */
            if (hundreds > 0 && tens == 0 && ones == 0)
            {
                if (hundreds == 1)
                    result += "etthundrade";
                else
                    result += OneThroughNineteen[hundreds] + "hundrade";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }

            /* > 100 */
            if (hundreds > 0)
            {
                result += OneThroughNineteen[hundreds] + "hundra";
            }

            /* exact tens */
            if (tens > 0 && ones == 0)
            {
                result += Tens[tens] + "nde";
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }

            /* > 20 */
            if (tens == 1)
            {
                result += OrdinalOneThroughNineteen[tens * 10 + ones];
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            else if (tens > 1)
            {
                result += Tens[tens] + OrdinalOneThroughNineteen[ones];
                ;
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
            else
            {
                result += OrdinalOneThroughNineteen[ones];
                return result.Substring(0, 1).ToUpper() + result.Substring(1);
            }
        }

        private static string NumberAsOrdinal(string languageCultureName, int levelNumber, string numFmt)
        {
            var levelAsString = levelNumber.ToString();

            if (levelAsString == null)
                return "";
            if (levelAsString.Trim() == "")
                return "";

            if (levelAsString.EndsWith("1"))
                return levelAsString + ":a";
            else if (levelAsString.EndsWith("2"))
                return levelAsString + ":a";
            else
                return levelAsString + ":e";
        }
    }
}
