using Clippit.Word.Enums;

namespace Clippit.Word.Extensions;

/// <summary>
/// Provides extension methods for the <see cref="NumberingFormatType"/> enumeration.
/// </summary>
public static class NumberingFormatTypeExtensions
{
    /// <summary>
    /// A case-insensitive mapping lookup dictionary to resolve OpenXML string format tokens
    /// to their strongly-typed <see cref="NumberingFormatType"/> equivalents.
    /// </summary>
    private static readonly Dictionary<string, NumberingFormatType> s_formatMap = new(StringComparer.OrdinalIgnoreCase)
    {
        { "none", NumberingFormatType.None },
        { "decimal", NumberingFormatType.Decimal },
        { "upperRoman", NumberingFormatType.UpperRoman },
        { "lowerRoman", NumberingFormatType.LowerRoman },
        { "upperLetter", NumberingFormatType.UpperLetter },
        { "lowerLetter", NumberingFormatType.LowerLetter },
        { "bullet", NumberingFormatType.Bullet },
        { "ordinal", NumberingFormatType.Ordinal },
        { "cardinalText", NumberingFormatType.CardinalText },
        { "ordinalText", NumberingFormatType.OrdinalText },
        { "decimalZero", NumberingFormatType.DecimalZero },
        { "decimalEnclosedCircle", NumberingFormatType.DecimalEnclosedCircle },
        { "ideographTraditional", NumberingFormatType.IdeographTraditional },
        { "chineseCounting", NumberingFormatType.ChineseCounting },
        { "chineseCountingThousand", NumberingFormatType.ChineseCountingThousand },
        { "01, 02, 03, ...", NumberingFormatType.DecimalPadded2 },
        { "001, 002, 003, ...", NumberingFormatType.DecimalPadded3 },
        { "0001, 0002, 0003, ...", NumberingFormatType.DecimalPadded4 },
        { "00001, 00002, 00003, ...", NumberingFormatType.DecimalPadded5 },
    };

    /// <summary>
    /// Parses a string representation of an OpenXML numbering format into its corresponding <see cref="NumberingFormatType"/> value.
    /// </summary>
    /// <param name="formatStr">The OpenXML format string to parse. Can be <see langword="null"/>.</param>
    /// <returns>
    /// The matching <see cref="NumberingFormatType"/> value; otherwise, <see cref="NumberingFormatType.Unspecified"/>
    /// if the string is <see langword="null"/>, empty, or unrecognized.
    /// </returns>
    public static NumberingFormatType ParseOpenXmlFormat(this string? formatStr)
    {
        if (formatStr is null)
        {
            return NumberingFormatType.Unspecified;
        }

        return s_formatMap.TryGetValue(formatStr, out var formatType) ? formatType : NumberingFormatType.Unspecified;
    }
}
