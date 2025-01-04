// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit
{
    internal class PartSHA1HashAnnotation(string hash)
    {
        public string Hash { get; set; } = hash;
    }
}
