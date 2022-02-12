// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Word
{
    public class ListItemTextGetter_zh_CN
    {
        public static string GetListItemText(string languageCultureName, int levelNumber, string numFmt)
        {
            var ccTDigitCharacters = new[] {
                "",
                "一",
                "二",
                "三",
                "四",
                "五",
                "六",
                "七",
                "八",
                "九",
            };
            var tenCharacter = "十";
            var hundredCharacter = "百";
            var thousandCharacter = "千";
            var andCharacter = "〇";

            var ccDigitCharacters = new[] {
                "○",
                "一",
                "二",
                "三",
                "四",
                "五",
                "六",
                "七",
                "八",
                "九",
            };

            var thousandsRemainder = levelNumber % 1000;
            var hundredsRemainder = levelNumber % 100;
            var thousands = levelNumber / 1000;
            var hundreds = (levelNumber % 1000) / 100;
            var tens = (levelNumber % 100) / 10;
            var ones = levelNumber % 10;

            if (numFmt == "chineseCounting")
            {
                return levelNumber switch
                {
                    >= 1 and <= 9 => ccDigitCharacters[levelNumber],
                    >= 10 and <= 19 when levelNumber == 10 => tenCharacter,
                    >= 10 and <= 19 => tenCharacter + ccDigitCharacters[ones],
                    >= 11 and <= 99 when ones == 0 => ccDigitCharacters[tens] + tenCharacter,
                    >= 11 and <= 99 => ccDigitCharacters[tens] + tenCharacter + ccDigitCharacters[ones],
                    >= 100 and <= 999 => ccDigitCharacters[hundreds] + ccDigitCharacters[tens] +
                                         ccDigitCharacters[ones],
                    >= 1000 and <= 9999 => ccDigitCharacters[thousands] + ccDigitCharacters[hundreds] +
                                           ccDigitCharacters[tens] + ccDigitCharacters[ones],
                    _ => levelNumber.ToString()
                };
            }
            if (numFmt == "chineseCountingThousand")
            {
                return levelNumber switch
                {
                    >= 1 and <= 9 => ccTDigitCharacters[levelNumber],
                    >= 10 and <= 19 => tenCharacter + ccTDigitCharacters[ones],
                    >= 20 and <= 99 => ccTDigitCharacters[tens] + tenCharacter + ccTDigitCharacters[ones],
                    >= 100 and <= 999 when hundredsRemainder == 0 => ccTDigitCharacters[hundreds] + hundredCharacter,
                    >= 100 and <= 999 when hundredsRemainder >= 1 && hundredsRemainder <= 9 => ccTDigitCharacters
                        [hundreds] + hundredCharacter + andCharacter + ccTDigitCharacters[levelNumber % 10],
                    >= 100 and <= 999 when ones == 0 => ccTDigitCharacters[hundreds] + hundredCharacter +
                                                        ccTDigitCharacters[tens] + tenCharacter,
                    >= 100 and <= 999 => ccTDigitCharacters[hundreds] + hundredCharacter + ccTDigitCharacters[tens] +
                                         tenCharacter + ccTDigitCharacters[ones],
                    >= 1000 and <= 9999 when thousandsRemainder == 0 => ccTDigitCharacters[thousands] +
                                                                        thousandCharacter,
                    >= 1000 and <= 9999 when thousandsRemainder >= 1 && thousandsRemainder <= 9 => ccTDigitCharacters
                            [thousands] + thousandCharacter + andCharacter +
                        GetListItemText("zh_CN", thousandsRemainder, numFmt),
                    >= 1000 and <= 9999 when thousandsRemainder >= 10 && thousandsRemainder <= 99 => ccTDigitCharacters
                            [thousands] + thousandCharacter + andCharacter + ccTDigitCharacters[tens] + tenCharacter +
                        ccTDigitCharacters[ones],
                    >= 1000 and <= 9999 when hundredsRemainder == 0 => ccTDigitCharacters[thousands] +
                                                                       thousandCharacter +
                                                                       ccTDigitCharacters[hundreds] + hundredCharacter,
                    >= 1000 and <= 9999 when hundredsRemainder >= 1 && hundredsRemainder <= 9 => ccTDigitCharacters
                            [thousands] + thousandCharacter + ccTDigitCharacters[hundreds] + hundredCharacter +
                        andCharacter + ccTDigitCharacters[ones],
                    >= 1000 and <= 9999 => ccTDigitCharacters[thousands] + thousandCharacter +
                                           ccTDigitCharacters[hundreds] + hundredCharacter + ccTDigitCharacters[tens] +
                                           tenCharacter + ccTDigitCharacters[ones],
                    _ => levelNumber.ToString()
                };
            }
            if (numFmt == "ideographTraditional")
            {
                var iDigitCharacters = new[] {
                    " ",
                    "甲",
                    "乙",
                    "丙",
                    "丁",
                    "戊",
                    "己",
                    "庚",
                    "辛",
                    "壬",
                    "癸",
                };
                if (levelNumber >= 1 && levelNumber <= 10)
                    return iDigitCharacters[levelNumber];
                return levelNumber.ToString();
            }
            return null;
        }
    }
}
