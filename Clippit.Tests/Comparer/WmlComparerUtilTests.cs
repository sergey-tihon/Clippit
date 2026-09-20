// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Security.Cryptography;
using System.Text;

namespace Clippit.Tests.Comparer;

/// <summary>
/// Unit tests for <see cref="WmlComparerUtil"/> covering hex encoding, SHA1 hashing
/// (both the short/stackalloc and long/heap code paths), and local-name to
/// <see cref="ComparisonUnitGroupType"/> mapping including the error case.
/// </summary>
public class WmlComparerUtilTests
{
    [Test]
    public async Task WU001_HexStringFromBytes_ByteArray_ReturnsLowercaseHex()
    {
        byte[] bytes = [0x00, 0x0f, 0xab, 0xff];

        var result = WmlComparerUtil.HexStringFromBytes(bytes);

        await Assert.That(result).IsEqualTo("000fabff");
    }

    [Test]
    public async Task WU002_HexStringFromBytes_ReadOnlySpan_ReturnsLowercaseHex()
    {
        ReadOnlySpan<byte> bytes = [0xde, 0xad, 0xbe, 0xef];

        var result = WmlComparerUtil.HexStringFromBytes(bytes);

        await Assert.That(result).IsEqualTo("deadbeef");
    }

    [Test]
    public async Task WU003_HexStringFromBytes_EmptyArray_ReturnsEmptyString()
    {
        var result = WmlComparerUtil.HexStringFromBytes([]);

        await Assert.That(result).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task WU004_HexStringFromBytes_NullArray_Throws()
    {
#pragma warning disable CS8625 // Cannot convert null literal to non-nullable reference type.
        await Assert.That(() => WmlComparerUtil.HexStringFromBytes((byte[])null)).Throws<ArgumentNullException>();
#pragma warning restore CS8625 // Cannot convert null literal to non-nullable reference type.
    }

    [Test]
    [Arguments("")]
    [Arguments("a")]
    [Arguments("hello world")]
    [Arguments("Привет 👋")]
    public async Task WU005_SHA1HashStringForUTF8String_ShortInput_MatchesExpectedHash(string input)
    {
        var expected = ComputeExpectedSha1Hex(Encoding.UTF8.GetBytes(input));

        var result = WmlComparerUtil.SHA1HashStringForUTF8String(input);

        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    public async Task WU006_SHA1HashStringForUTF8String_LongInput_MatchesExpectedHash()
    {
        // Exceeds MaxStackallocUtf8Bytes (4096) to exercise the heap-allocated code path.
        var input = new string('x', 5000);
        var expected = ComputeExpectedSha1Hex(Encoding.UTF8.GetBytes(input));

        var result = WmlComparerUtil.SHA1HashStringForUTF8String(input);

        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    public async Task WU007_SHA1HashStringForByteArray_ReturnsExpectedHash()
    {
        var bytes = Encoding.UTF8.GetBytes("some content to hash");
        var expected = ComputeExpectedSha1Hex(bytes);

        var result = WmlComparerUtil.SHA1HashStringForByteArray(bytes);

        await Assert.That(result).IsEqualTo(expected);
    }

    [Test]
    public async Task WU008_SHA1HashStringForByteArray_NullArray_Throws()
    {
#pragma warning disable CS8625 // Cannot convert null literal to non-nullable reference type.
        await Assert.That(() => WmlComparerUtil.SHA1HashStringForByteArray(null)).Throws<ArgumentNullException>();
#pragma warning restore CS8625 // Cannot convert null literal to non-nullable reference type.
    }

    [Test]
    [Arguments("p", "Paragraph")]
    [Arguments("tbl", "Table")]
    [Arguments("tr", "Row")]
    [Arguments("tc", "Cell")]
    [Arguments("txbxContent", "Textbox")]
    public async Task WU009_ComparisonUnitGroupTypeFromLocalName_KnownNames_ReturnsExpectedType(
        string localName,
        string expectedName
    )
    {
        var result = WmlComparerUtil.ComparisonUnitGroupTypeFromLocalName(localName);

        await Assert.That(result.ToString()).IsEqualTo(expectedName);
    }

    [Test]
    public async Task WU010_ComparisonUnitGroupTypeFromLocalName_UnknownName_Throws()
    {
        await Assert
            .That(() => WmlComparerUtil.ComparisonUnitGroupTypeFromLocalName("unknown"))
            .Throws<ArgumentOutOfRangeException>();
    }

    private static string ComputeExpectedSha1Hex(byte[] bytes) => Convert.ToHexStringLower(SHA1.HashData(bytes));
}
