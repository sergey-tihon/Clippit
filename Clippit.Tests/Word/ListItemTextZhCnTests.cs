// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Word;

namespace Clippit.Tests.Word;

/// <summary>
/// Unit tests for <see cref="ListItemTextGetter_zh_CN.GetListItemText"/>.
/// </summary>
public class ListItemTextZhCnTests
{
    // ── chineseCounting — ones ────────────────────────────────────────────────

    [Test]
    [Arguments(1, "一")]
    [Arguments(5, "五")]
    [Arguments(9, "九")]
    public async Task LZhCn001_ChineseCounting_Ones_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "chineseCounting");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── chineseCounting — tens ────────────────────────────────────────────────

    [Test]
    [Arguments(10, "十")]
    [Arguments(11, "十一")]
    [Arguments(19, "十九")]
    [Arguments(20, "二十")]
    [Arguments(21, "二十一")]
    [Arguments(99, "九十九")]
    public async Task LZhCn002_ChineseCounting_Tens_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "chineseCounting");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── chineseCounting — hundreds / thousands ────────────────────────────────

    [Test]
    [Arguments(100, "一○○")]
    [Arguments(101, "一○一")]
    [Arguments(123, "一二三")]
    [Arguments(1000, "一○○○")]
    [Arguments(1234, "一二三四")]
    public async Task LZhCn003_ChineseCounting_HundredsAndThousands_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "chineseCounting");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── chineseCountingThousand — ones / teens ────────────────────────────────

    [Test]
    [Arguments(1, "一")]
    [Arguments(9, "九")]
    [Arguments(10, "十")]
    [Arguments(11, "十一")]
    [Arguments(19, "十九")]
    public async Task LZhCn004_ChineseCountingThousand_OnesToNineteen_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "chineseCountingThousand");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── chineseCountingThousand — tens ────────────────────────────────────────

    [Test]
    [Arguments(20, "二十")]
    [Arguments(21, "二十一")]
    [Arguments(99, "九十九")]
    public async Task LZhCn005_ChineseCountingThousand_Tens_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "chineseCountingThousand");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── chineseCountingThousand — hundreds ────────────────────────────────────

    [Test]
    [Arguments(100, "一百")]
    [Arguments(101, "一百〇一")]
    [Arguments(110, "一百一十")]
    [Arguments(111, "一百一十一")]
    [Arguments(200, "二百")]
    public async Task LZhCn006_ChineseCountingThousand_Hundreds_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "chineseCountingThousand");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── chineseCountingThousand — thousands ───────────────────────────────────

    [Test]
    [Arguments(1000, "一千")]
    [Arguments(1001, "一千〇一")]
    [Arguments(1010, "一千〇一十")]
    [Arguments(1100, "一千一百")]
    [Arguments(1101, "一千一百〇一")]
    [Arguments(1234, "一千二百三十四")]
    [Arguments(2000, "二千")]
    public async Task LZhCn007_ChineseCountingThousand_Thousands_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "chineseCountingThousand");
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── ideographTraditional ─────────────────────────────────────────────────

    [Test]
    [Arguments(1, "甲")]
    [Arguments(2, "乙")]
    [Arguments(3, "丙")]
    [Arguments(4, "丁")]
    [Arguments(5, "戊")]
    [Arguments(6, "己")]
    [Arguments(7, "庚")]
    [Arguments(8, "辛")]
    [Arguments(9, "壬")]
    [Arguments(10, "癸")]
    public async Task LZhCn008_IdeographTraditional_OneToTen_ReturnsExpected(int number, string expected)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "ideographTraditional");
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(11)]
    [Arguments(100)]
    public async Task LZhCn009_IdeographTraditional_OutOfRange_FallsBackToDecimal(int number)
    {
        var result = ListItemTextGetter_zh_CN.GetListItemText("zh-CN", number, "ideographTraditional");
        await Assert.That(result).IsEqualTo(number.ToString());
    }
}
