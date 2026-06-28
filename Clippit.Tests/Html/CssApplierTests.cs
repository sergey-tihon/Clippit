// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Reflection;
using Clippit.Html;

namespace Clippit.Tests.Html;

public class CssApplierTests : TestsBase
{
    [Test]
    [Arguments("#first { color: red; }", 1)]
    [Arguments("#first#second { color: red; }", 2)]
    [Arguments("div#first#second { color: red; }", 2)]
    public async Task HW005_CountIdAttributesInSelectorIncludesChildIds(string css, int expectedCount)
    {
        var parser = new CssParser();
        var selector = parser.ParseText(css).RuleSets.Single().Selectors.Single();
        var method = typeof(CssApplier).GetMethod(
            "CountIdAttributesInSelector",
            BindingFlags.NonPublic | BindingFlags.Static
        );

        await Assert.That(method).IsNotNull();

        var count = (int)method!.Invoke(null, new object[] { selector })!;
        await Assert.That(count).IsEqualTo(expectedCount);
    }
}
