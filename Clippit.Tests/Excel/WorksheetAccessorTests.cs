// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Excel;

namespace Clippit.Tests.Excel;

/// <summary>
/// Unit tests for <see cref="WorksheetAccessor.GetColumnId"/> and
/// <see cref="WorksheetAccessor.GetColumnNumber"/> — the 1-based column-address
/// utility methods used throughout spreadsheet manipulation.
/// </summary>
public class WorksheetAccessorTests
{
    // ── GetColumnId ──────────────────────────────────────────────────────────
    // Converts a 1-based column number to its Excel column-letter identifier.

    [Test]
    [Arguments(1, "A")]
    [Arguments(2, "B")]
    [Arguments(26, "Z")]
    [Arguments(27, "AA")]
    [Arguments(28, "AB")]
    [Arguments(52, "AZ")]
    [Arguments(53, "BA")]
    [Arguments(702, "ZZ")]
    [Arguments(703, "AAA")]
    [Arguments(704, "AAB")]
    [Arguments(18278, "ZZZ")]
    public async Task WA001_GetColumnId_KnownValues_ReturnsExpectedLetter(int column, string expected)
    {
        var result = WorksheetAccessor.GetColumnId(column);
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── GetColumnNumber ──────────────────────────────────────────────────────
    // Extracts the 1-based column number from a cell reference such as "A1" or "AA99".
    // Digit characters (the row number) are ignored; only letters are used.

    [Test]
    [Arguments("A1", 1)]
    [Arguments("B1", 2)]
    [Arguments("Z1", 26)]
    [Arguments("AA1", 27)]
    [Arguments("AB1", 28)]
    [Arguments("AZ1", 52)]
    [Arguments("BA1", 53)]
    [Arguments("ZZ999", 702)]
    [Arguments("AAA1", 703)]
    [Arguments("ZZZ1048576", 18278)]
    public async Task WA002_GetColumnNumber_CellReference_ReturnsExpected1BasedColumn(
        string cellReference,
        int expected
    )
    {
        var result = WorksheetAccessor.GetColumnNumber(cellReference);
        await Assert.That(result).IsEqualTo(expected);
    }

    // ── Round-trip: GetColumnId → GetColumnNumber ────────────────────────────

    [Test]
    [Arguments(1)]
    [Arguments(26)]
    [Arguments(27)]
    [Arguments(52)]
    [Arguments(702)]
    [Arguments(703)]
    [Arguments(18278)]
    public async Task WA003_RoundTrip_ColumnId_ThenColumnNumber_IsIdentity(int originalColumn)
    {
        var address = WorksheetAccessor.GetColumnId(originalColumn) + "1";
        var recovered = WorksheetAccessor.GetColumnNumber(address);
        await Assert.That(recovered).IsEqualTo(originalColumn);
    }
}
