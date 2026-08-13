// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Drawing;
using Clippit.Internal;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="ColorParser"/> covering valid names, unknown names,
/// and the null-input edge case that must not throw.
/// </summary>
public class ColorParserTests
{
    [Test]
    public async Task CP001_TryFromName_KnownColor_ReturnsTrueAndColor()
    {
        var result = ColorParser.TryFromName("Red", out var color);

        await Assert.That(result).IsTrue();
        await Assert.That(color).IsEqualTo(Color.Red);
    }

    [Test]
    public async Task CP002_TryFromName_UnknownName_ReturnsTrueButColorIsEmpty()
    {
        // System.Drawing.Color.FromName treats any non-null string as a "named" color
        // (IsNamedColor is true even for unrecognized names), but the resulting color
        // has no ARGB value set (it is the default/empty Color).
        var result = ColorParser.TryFromName("NotARealColorName", out var color);

        await Assert.That(result).IsTrue();
        await Assert.That(color.IsKnownColor).IsFalse();
        await Assert.That(color.ToArgb()).IsEqualTo(0);
    }

    [Test]
    public async Task CP003_TryFromName_NullName_DoesNotThrowAndReturnsFalse()
    {
        var result = ColorParser.TryFromName(null!, out var color);

        await Assert.That(result).IsFalse();
        await Assert.That(color).IsEqualTo(default(Color));
    }

    [Test]
    public async Task CP004_IsValidName_KnownColor_ReturnsTrue()
    {
        await Assert.That(ColorParser.IsValidName("Blue")).IsTrue();
    }

    [Test]
    public async Task CP005_IsValidName_EmptyString_ReturnsTrueForEmptyNamedColor()
    {
        // Empty string is still accepted by Color.FromName as a "named" color (IsNamedColor true),
        // matching the behavior verified in CP002 for unrecognized names.
        await Assert.That(ColorParser.IsValidName("")).IsTrue();
    }

    [Test]
    public async Task CP006_FromName_KnownColor_ReturnsExpectedColor()
    {
        var color = ColorParser.FromName("Green");

        await Assert.That(color).IsEqualTo(Color.Green);
    }
}
