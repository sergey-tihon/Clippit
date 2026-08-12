using System.Diagnostics.CodeAnalysis;

namespace Clippit.Word.Enums;

[SuppressMessage("Naming", "CA1720:Identifier contains type name", Justification = "Matches OpenXML w:numFmt spec")]
public enum NumberingFormatType
{
    None = 0,
    Decimal = 1,
    UpperRoman = 2,
    LowerRoman = 3,
    UpperLetter = 4,
    LowerLetter = 5,
    Bullet = 6,
    Ordinal = 7,
    CardinalText = 8,
    OrdinalText = 9,
    DecimalZero = 10,
    DecimalEnclosedCircle = 11,
    IdeographTraditional = 12,
    ChineseCounting = 13,
    ChineseCountingThousand = 14,
    /// <summary>
    /// "01, 02, 03, ..."
    /// </summary>
    DecimalPadded2 = 15,
    /// <summary>
    /// "001, 002, 003, ..."
    /// </summary>
    DecimalPadded3 = 16,
    /// <summary>
    /// "0001, 0002, 0003, ..."
    /// </summary>
    DecimalPadded4 = 17,
    /// <summary>
    /// "00001, 00002, 00003, ..."
    /// </summary>
    DecimalPadded5 = 18
}
