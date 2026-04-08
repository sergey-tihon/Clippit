// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Excel;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Excel;

/// <summary>
/// Unit tests for the static utility helpers on <see cref="XlsxTables"/>:
/// <list type="bullet">
/// <item><description><see cref="XlsxTables.SplitAddress"/> — splits a cell address (e.g. "AB12") into column and row parts</description></item>
/// <item><description><see cref="XlsxTables.IndexToColumnAddress"/> — converts a 0-based column index to a column letter string</description></item>
/// <item><description><see cref="XlsxTables.ColumnAddressToIndex"/> — converts a column letter string to a 0-based column index</description></item>
/// </list>
/// These helpers are critical for spreadsheet data access and have no other dedicated tests.
/// </summary>
public class XlsxTablesUtilTests
{
    // ── SplitAddress ────────────────────────────────────────────────────────

    [Test]
    [Arguments("A1", "A", "1")]
    [Arguments("Z99", "Z", "99")]
    [Arguments("AA10", "AA", "10")]
    [Arguments("AAA1", "AAA", "1")]
    [Arguments("XFD1048576", "XFD", "1048576")]
    public async Task XT001_SplitAddress_ValidAddress_ReturnsTwoParts(
        string address,
        string expectedCol,
        string expectedRow
    )
    {
        var parts = XlsxTables.SplitAddress(address);
        await Assert.That(parts).Count().IsEqualTo(2);
        await Assert.That(parts[0]).IsEqualTo(expectedCol);
        await Assert.That(parts[1]).IsEqualTo(expectedRow);
    }

    [Test]
    [Arguments("A")]
    [Arguments("ABC")]
    public async Task XT002_SplitAddress_NoRowNumber_ThrowsFileFormatException(string address)
    {
        await Assert.That(() => XlsxTables.SplitAddress(address)).Throws<FileFormatException>();
    }

    // ── IndexToColumnAddress ────────────────────────────────────────────────

    [Test]
    [Arguments(0, "A")]
    [Arguments(1, "B")]
    [Arguments(25, "Z")]
    [Arguments(26, "AA")]
    [Arguments(27, "AB")]
    [Arguments(51, "AZ")]
    [Arguments(52, "BA")]
    [Arguments(701, "ZZ")]
    [Arguments(702, "AAA")]
    [Arguments(703, "AAB")]
    [Arguments(18277, "ZZZ")]
    public async Task XT003_IndexToColumnAddress_KnownValues_ReturnsExpectedAddress(int index, string expected)
    {
        var result = XlsxTables.IndexToColumnAddress(index);
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments(18278)]
    [Arguments(99999)]
    public async Task XT004_IndexToColumnAddress_OutOfRange_Throws(int index)
    {
        await Assert.That(() => XlsxTables.IndexToColumnAddress(index)).Throws<Exception>();
    }

    // ── ColumnAddressToIndex ────────────────────────────────────────────────

    [Test]
    [Arguments("A", 0)]
    [Arguments("B", 1)]
    [Arguments("Z", 25)]
    [Arguments("AA", 26)]
    [Arguments("AB", 27)]
    [Arguments("AZ", 51)]
    [Arguments("BA", 52)]
    [Arguments("ZZ", 701)]
    [Arguments("AAA", 702)]
    [Arguments("AAB", 703)]
    [Arguments("ZZZ", 18277)]
    public async Task XT005_ColumnAddressToIndex_KnownValues_ReturnsExpectedIndex(string address, int expected)
    {
        var result = XlsxTables.ColumnAddressToIndex(address);
        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    [Arguments("AAAA")]
    [Arguments("ABCDE")]
    public async Task XT006_ColumnAddressToIndex_TooLong_Throws(string address)
    {
        await Assert.That(() => XlsxTables.ColumnAddressToIndex(address)).Throws<FileFormatException>();
    }

    // ── Round-trip: IndexToColumnAddress / ColumnAddressToIndex ────────────

    [Test]
    [Arguments(0)]
    [Arguments(1)]
    [Arguments(25)]
    [Arguments(26)]
    [Arguments(51)]
    [Arguments(701)]
    [Arguments(702)]
    [Arguments(18277)]
    public async Task XT007_RoundTrip_IndexToAddress_ThenAddressToIndex_IsIdentity(int originalIndex)
    {
        var address = XlsxTables.IndexToColumnAddress(originalIndex);
        var recoveredIndex = XlsxTables.ColumnAddressToIndex(address);
        await Assert.That(recoveredIndex).IsEqualTo(originalIndex);
    }

    [Test]
    [Arguments("A")]
    [Arguments("Z")]
    [Arguments("AA")]
    [Arguments("AZ")]
    [Arguments("BA")]
    [Arguments("ZZ")]
    [Arguments("AAA")]
    [Arguments("ZZZ")]
    public async Task XT008_RoundTrip_AddressToIndex_ThenIndexToAddress_IsIdentity(string originalAddress)
    {
        var index = XlsxTables.ColumnAddressToIndex(originalAddress);
        var recoveredAddress = XlsxTables.IndexToColumnAddress(index);
        await Assert.That(recoveredAddress).IsEqualTo(originalAddress);
    }
}
