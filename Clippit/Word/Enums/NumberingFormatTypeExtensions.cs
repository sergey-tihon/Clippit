namespace Clippit.Word.Enums;

public static class NumberingFormatTypeExtensions
{
    public static NumberingFormatType ParseOpenXmlFormat(string formatStr)
    {
        return formatStr switch
        {
            "decimal" => NumberingFormatType.Decimal,
            "upperRoman" => NumberingFormatType.UpperRoman,
            "lowerRoman" => NumberingFormatType.LowerRoman,
            "upperLetter" => NumberingFormatType.UpperLetter,
            "lowerLetter" => NumberingFormatType.LowerLetter,
            "bullet" => NumberingFormatType.Bullet,
            "ordinal" => NumberingFormatType.Ordinal,
            "cardinalText" => NumberingFormatType.CardinalText,
            "ordinalText" => NumberingFormatType.OrdinalText,
            "decimalZero" => NumberingFormatType.DecimalZero,
            "decimalEnclosedCircle" => NumberingFormatType.DecimalEnclosedCircle,
            "ideographTraditional" => NumberingFormatType.IdeographTraditional,
            "chineseCounting" => NumberingFormatType.ChineseCounting,
            "chineseCountingThousand" => NumberingFormatType.ChineseCountingThousand,
            "01, 02, 03, ..." => NumberingFormatType.DecimalPadded2,
            "001, 002, 003, ..." => NumberingFormatType.DecimalPadded3,
            "0001, 0002, 0003, ..." => NumberingFormatType.DecimalPadded4,
            "00001, 00002, 00003, ..." => NumberingFormatType.DecimalPadded5,
            _ => NumberingFormatType.None
        };
    }
}
