// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Security.Cryptography;
using System.Text;

namespace Clippit
{
    internal static class WmlComparerUtil
    {
        public static string SHA1HashStringForUTF8String(string s)
        {
            var bytes = Encoding.UTF8.GetBytes(s);
            using var sha1 = SHA1.Create();
            var hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string SHA1HashStringForByteArray(byte[] bytes)
        {
            using var sha1 = SHA1.Create();
            var hashBytes = sha1.ComputeHash(bytes);
            return HexStringFromBytes(hashBytes);
        }

        public static string HexStringFromBytes(byte[] bytes)
        {
            var sb = new StringBuilder();
            foreach (var b in bytes)
            {
                var hex = b.ToString("x2");
                sb.Append(hex);
            }

            return sb.ToString();
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
