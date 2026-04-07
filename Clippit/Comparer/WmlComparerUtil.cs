// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Security.Cryptography;
using System.Text;

namespace Clippit
{
    internal static class WmlComparerUtil
    {
        public static string SHA1HashStringForUTF8String(string s)
        {
            var bytes = Encoding.UTF8.GetBytes(s);
            var hashBytes = SHA1.HashData(bytes);
            return Convert.ToHexString(hashBytes).ToLowerInvariant();
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            var hashBytes = SHA1.HashData(bytes);
            return Convert.ToHexString(hashBytes).ToLowerInvariant();
        }

        public static string HexStringFromBytes(byte[] bytes) =>
            Convert.ToHexString(bytes).ToLowerInvariant();

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
