using Clippit.Word.Enums;
using Clippit.Word.Extensions;

namespace Clippit.Tests.Word.Extensions;

public sealed class NumberingFormatTypeExtensionsTests
{
    [Test]
    [Arguments("none", NumberingFormatType.None)]
    [Arguments("decimal", NumberingFormatType.Decimal)]
    [Arguments("upperRoman", NumberingFormatType.UpperRoman)]
    [Arguments("lowerRoman", NumberingFormatType.LowerRoman)]
    [Arguments("upperLetter", NumberingFormatType.UpperLetter)]
    [Arguments("lowerLetter", NumberingFormatType.LowerLetter)]
    [Arguments("bullet", NumberingFormatType.Bullet)]
    [Arguments("ordinal", NumberingFormatType.Ordinal)]
    [Arguments("cardinalText", NumberingFormatType.CardinalText)]
    [Arguments("ordinalText", NumberingFormatType.OrdinalText)]
    [Arguments("decimalZero", NumberingFormatType.DecimalZero)]
    [Arguments("decimalEnclosedCircle", NumberingFormatType.DecimalEnclosedCircle)]
    [Arguments("ideographTraditional", NumberingFormatType.IdeographTraditional)]
    [Arguments("chineseCounting", NumberingFormatType.ChineseCounting)]
    [Arguments("chineseCountingThousand", NumberingFormatType.ChineseCountingThousand)]
    [Arguments("01, 02, 03, ...", NumberingFormatType.DecimalPadded2)]
    [Arguments("001, 002, 003, ...", NumberingFormatType.DecimalPadded3)]
    [Arguments("0001, 0002, 0003, ...", NumberingFormatType.DecimalPadded4)]
    [Arguments("00001, 00002, 00003, ...", NumberingFormatType.DecimalPadded5)]
    public async Task ParseOpenXmlFormat_WithValidString_ReturnsExpectedFormatType(string input, NumberingFormatType expected)
    {
        // Act
        var result = input.ParseOpenXmlFormat();

        // Assert
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(null)]
    [Arguments("")]
    [Arguments("   ")]
    [Arguments("unknownFormat")]
    public async Task ParseOpenXmlFormat_WithInvalidOrNullString_ReturnsUnspecified(string? input)
    {
        // Act
        var result = input.ParseOpenXmlFormat();

        // Assert
        await Assert.That(result).IsEqualTo(NumberingFormatType.Unspecified);
    }

    [Test]
    [Arguments("Decimal")]
    [Arguments("ORDINAL")]
    [Arguments("UPPERROMAN")]
    [Arguments("LowerRoman")]
    [Arguments("BULLET")]
    [Arguments("ordinaltext")]
    [Arguments("DecimalZero")]
    [Arguments("ChineseCountingThousand")]
    public async Task ParseOpenXmlFormat_IsCaseSensitive_ReturnsUnspecifiedWhenCaseMismatched(string input)
    {
        // Act
        var result = input.ParseOpenXmlFormat();

        // Assert
        await Assert.That(result).IsEqualTo(NumberingFormatType.Unspecified);
    }
}
