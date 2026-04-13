// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Security.Cryptography;
using System.Text;

namespace Clippit
{
    internal static class WmlComparerUtil
    {
        public static string HexStringFromBytes(byte[] bytes) =>
#if NET9_0_OR_GREATER
            Convert.ToHexStringLower(bytes);
#else
            Convert.ToHexString(bytes).ToLowerInvariant();
#endif

        // Hot path in WmlComparer: called for every comparison unit atom.
        // Uses stackalloc for both the UTF-8 byte buffer and the 20-byte SHA-1 output
        // to avoid heap allocations for the common case of short strings.
        public static string SHA1HashStringForUTF8String(string s)
        {
            var maxByteCount = Encoding.UTF8.GetMaxByteCount(s.Length);
            if (maxByteCount <= 4096)
            {
                Span<byte> inputBuffer = stackalloc byte[maxByteCount];
                var actualBytes = Encoding.UTF8.GetBytes(s, inputBuffer);
                Span<byte> hashBuffer = stackalloc byte[SHA1.HashSizeInBytes];
                SHA1.TryHashData(inputBuffer[..actualBytes], hashBuffer, out _);
#if NET9_0_OR_GREATER
                return Convert.ToHexStringLower(hashBuffer);
#else
                return Convert.ToHexString(hashBuffer).ToLowerInvariant();
#endif
            }
            return HexStringFromBytes(SHA1.HashData(Encoding.UTF8.GetBytes(s)));
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            Span<byte> hashBuffer = stackalloc byte[SHA1.HashSizeInBytes];
            SHA1.TryHashData(bytes, hashBuffer, out _);
#if NET9_0_OR_GREATER
            return Convert.ToHexStringLower(hashBuffer);
#else
            return Convert.ToHexString(hashBuffer).ToLowerInvariant();
#endif
        }

        public static ComparisonUnitGroupType ComparisonUnitGroupTypeFromLocalName(string localName) =>
            localName switch
            {
                "p" => ComparisonUnitGroupType.Paragraph,
                "tbl" => ComparisonUnitGroupType.Table,
                "tr" => ComparisonUnitGroupType.Row,
                "tc" => ComparisonUnitGroupType.Cell,
                "txbxContent" => ComparisonUnitGroupType.Textbox,
                _ => throw new ArgumentOutOfRangeException(
                    nameof(localName),
                    $@"Unsupported localName: '{localName}'."
                ),
            };
    }
}
