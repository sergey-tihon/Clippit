// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Security.Cryptography;
using System.Text;

namespace Clippit
{
    internal static class WmlComparerUtil
    {
        // Maximum UTF-8 byte count for which we use stackalloc to avoid heap allocations
        // for the input buffer. In the worst case, 4096 bytes is only about 1024 UTF-8
        // characters, and the actual stackalloc check uses Encoding.UTF8.GetMaxByteCount.
        private const int MaxStackallocUtf8Bytes = 4096;

        public static string HexStringFromBytes(byte[] bytes) => HexStringFromBytes((ReadOnlySpan<byte>)bytes);

        public static string HexStringFromBytes(ReadOnlySpan<byte> bytes) =>
#if NET9_0_OR_GREATER
            Convert.ToHexStringLower(bytes);
#else
            Convert.ToHexString(bytes).ToLowerInvariant();
#endif

        // Hot path in WmlComparer: called for every comparison unit atom.
        // Uses stackalloc for both the UTF-8 input buffer and the 20-byte SHA-1 output buffer
        // to avoid heap-allocating those temporary byte[] buffers for the common case of
        // short strings. The returned hex string is still allocated.
        public static string SHA1HashStringForUTF8String(string s)
        {
            var maxByteCount = Encoding.UTF8.GetMaxByteCount(s.Length);
            if (maxByteCount <= MaxStackallocUtf8Bytes)
            {
                Span<byte> inputBuffer = stackalloc byte[maxByteCount];
                var actualBytes = Encoding.UTF8.GetBytes(s, inputBuffer);
                Span<byte> hashBuffer = stackalloc byte[SHA1.HashSizeInBytes];
                if (!SHA1.TryHashData(inputBuffer[..actualBytes], hashBuffer, out _))
                    throw new CryptographicException("SHA1.TryHashData failed unexpectedly.");
                return HexStringFromBytes((ReadOnlySpan<byte>)hashBuffer);
            }

            var utf8Bytes = Encoding.UTF8.GetBytes(s);
            Span<byte> longStringHashBuffer = stackalloc byte[SHA1.HashSizeInBytes];
            if (!SHA1.TryHashData(utf8Bytes, longStringHashBuffer, out _))
                throw new CryptographicException("SHA1.TryHashData failed unexpectedly.");
            return HexStringFromBytes((ReadOnlySpan<byte>)longStringHashBuffer);
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            Span<byte> hashBuffer = stackalloc byte[SHA1.HashSizeInBytes];
            if (!SHA1.TryHashData(bytes, hashBuffer, out _))
                throw new CryptographicException("SHA1.TryHashData failed unexpectedly.");
            return HexStringFromBytes((ReadOnlySpan<byte>)hashBuffer);
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
