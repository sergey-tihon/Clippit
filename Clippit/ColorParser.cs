// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Drawing;

namespace Clippit
{
    public static class ColorParser
    {
        public static Color FromName(string name) => 
            Color.FromName(name);

        public static bool TryFromName(string name, out Color color)
        {
            try
            {
                color = Color.FromName(name);

                return color.IsNamedColor;
            }
            catch
            {
                color = default;

                return false;
            }
        }

        public static bool IsValidName(string name) => 
            TryFromName(name, out _);
    }
}
