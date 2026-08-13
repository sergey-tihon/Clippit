using System.Diagnostics.CodeAnalysis;

namespace Clippit.Word.Enums;

/// <summary>
/// Specifies the formatting type applied to a numbering sequence.
/// </summary>
/// <remarks>These values correspond to the OpenXML <c>w:numFmt</c> specification.</remarks>
[SuppressMessage("Naming", "CA1720:Identifier contains type name", Justification = "Matches OpenXML w:numFmt spec")]
public enum NumberingFormatType
{
    /// <summary>
    /// The numbering format is unspecified, unrecognized, or missing.
    /// </summary>
    Unspecified = 0,

    /// <summary>
    /// Explicitly specifies that no numbering format is applied.
    /// </summary>
    None = 1,

    /// <summary>
    /// Decimal numbers.
    /// </summary>
    /// <example>1, 2, 3, ...</example>
    Decimal = 2,

    /// <summary>
    /// Uppercase Roman numerals.
    /// </summary>
    /// <example>I, II, III, ...</example>
    UpperRoman = 3,

    /// <summary>
    /// Lowercase Roman numerals.
    /// </summary>
    /// <example>i, ii, iii, ...</example>
    LowerRoman = 4,

    /// <summary>
    /// Uppercase Latin letters.
    /// </summary>
    /// <example>A, B, C, ...</example>
    UpperLetter = 5,

    /// <summary>
    /// Lowercase Latin letters.
    /// </summary>
    /// <example>a, b, c, ...</example>
    LowerLetter = 6,

    /// <summary>
    /// Bullet characters.
    /// </summary>
    Bullet = 7,

    /// <summary>
    /// Ordinal numbers.
    /// </summary>
    /// <example>1st, 2nd, 3rd, ...</example>
    Ordinal = 8,

    /// <summary>
    /// Cardinal text numbers.
    /// </summary>
    /// <example>One, Two, Three, ...</example>
    CardinalText = 9,

    /// <summary>
    /// Ordinal text numbers.
    /// </summary>
    /// <example>First, Second, Third, ...</example>
    OrdinalText = 10,

    /// <summary>
    /// Decimal numbers.
    /// </summary>
    /// <example>1, 2, 3, ...</example>
    DecimalZero = 11,

    /// <summary>
    /// Decimal numbers enclosed in a circle.
    /// </summary>
    /// <example>①, ②, ③, ...</example>
    DecimalEnclosedCircle = 12,

    /// <summary>
    /// Traditional ideograph numbering.
    /// </summary>
    IdeographTraditional = 13,

    /// <summary>
    /// Chinese counting numbers.
    /// </summary>
    ChineseCounting = 14,

    /// <summary>
    /// Chinese counting thousand numbers.
    /// </summary>
    ChineseCountingThousand = 15,

    /// <summary>
    /// Decimal numbers padded with leading zeros to a minimum length of two characters.
    /// </summary>
    /// <example>01, 02, 03, ...</example>
    DecimalPadded2 = 16,

    /// <summary>
    /// Decimal numbers padded with leading zeros to a minimum length of three characters.
    /// </summary>
    /// <example>001, 002, 003, ...</example>
    DecimalPadded3 = 17,

    /// <summary>
    /// Decimal numbers padded with leading zeros to a minimum length of four characters.
    /// </summary>
    /// <example>0001, 0002, 0003, ...</example>
    DecimalPadded4 = 18,

    /// <summary>
    /// Decimal numbers padded with leading zeros to a minimum length of five characters.
    /// </summary>
    /// <example>00001, 00002, 00003, ...</example>
    DecimalPadded5 = 19
}
