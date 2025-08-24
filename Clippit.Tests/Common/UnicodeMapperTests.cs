// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
/***************************************************************************

Copyright (c) Microsoft Corporation 2016.

This code is licensed using the Microsoft Public License (Ms-PL).  The text of the license can be found here:

http://www.microsoft.com/resources/sharedsource/licensingbasics/publiclicense.mspx

Developer: Thomas Barnekow
Email: thomas@barnekow.info

***************************************************************************/
using System.Xml.Linq;

namespace Clippit.Tests.Common;

public class UnicodeMapperTests
{
    [Test]
    public async Task CanStringifyRunAndTextElements()
    {
        const string TextValue = "Hello World!";
        var textElement = new XElement(W.t, TextValue);
        var runElement = new XElement(W.r, textElement);
        var formattedRunElement = new XElement(W.r, new XElement(W.rPr, new XElement(W.b)), textElement);
        await Assert.That(UnicodeMapper.RunToString(textElement)).IsEqualTo(TextValue);
        await Assert.That(UnicodeMapper.RunToString(runElement)).IsEqualTo(TextValue);
        await Assert.That(UnicodeMapper.RunToString(formattedRunElement)).IsEqualTo(TextValue);
    }

    [Test]
    public async Task CanStringifySpecialElements()
    {
        await Assert
            .That(UnicodeMapper.RunToString(new XElement(W.cr)).First())
            .IsEqualTo(UnicodeMapper.CarriageReturn);
        await Assert
            .That(UnicodeMapper.RunToString(new XElement(W.br)).First())
            .IsEqualTo(UnicodeMapper.CarriageReturn);
        await Assert
            .That(UnicodeMapper.RunToString(new XElement(W.br, new XAttribute(W.type, "page"))).First())
            .IsEqualTo(UnicodeMapper.FormFeed);
        await Assert
            .That(UnicodeMapper.RunToString(new XElement(W.noBreakHyphen)).First())
            .IsEqualTo(UnicodeMapper.NonBreakingHyphen);
        await Assert
            .That(UnicodeMapper.RunToString(new XElement(W.softHyphen)).First())
            .IsEqualTo(UnicodeMapper.SoftHyphen);
        await Assert.That(UnicodeMapper.RunToString(new XElement(W.tab)).First()).IsEqualTo(UnicodeMapper.SoftHyphen);
    }

    [Test]
    public async Task CanCreateRunChildElementsFromSpecialCharacters()
    {
        await Assert.That(UnicodeMapper.CharToRunChild(UnicodeMapper.CarriageReturn).Name).IsEqualTo(W.br);
        await Assert.That(UnicodeMapper.CharToRunChild(UnicodeMapper.NonBreakingHyphen).Name).IsEqualTo(W.softHyphen);
        await Assert.That(UnicodeMapper.CharToRunChild(UnicodeMapper.SoftHyphen).Name).IsEqualTo(W.softHyphen);
        await Assert.That(UnicodeMapper.CharToRunChild(UnicodeMapper.HorizontalTabulation).Name).IsEqualTo(W.tab);
        var element = UnicodeMapper.CharToRunChild(UnicodeMapper.FormFeed);
        await Assert.That(element.Name).IsEqualTo(W.br);
        await Assert.That(element.Attribute(W.type).Value).IsEqualTo("page");
        await Assert.That(UnicodeMapper.CharToRunChild('\r').Name).IsEqualTo(W.br);
    }

    [Test]
    public async Task CanCreateCoalescedRuns()
    {
        const string TextString = "This is only text.";
        const string MixedString = "First\tSecond\tThird";
        var textRuns = UnicodeMapper.StringToCoalescedRunList(TextString, null);
        var mixedRuns = UnicodeMapper.StringToCoalescedRunList(MixedString, null);
        await Assert.That(textRuns).HasSingleItem();
        await Assert.That(mixedRuns).HasCount(5);
        await Assert.That(mixedRuns.Elements(W.t).Skip(0).First().Value).IsEqualTo("First");
        await Assert.That(mixedRuns.Elements(W.t).Skip(1).First().Value).IsEqualTo("Second");
        await Assert.That(mixedRuns.Elements(W.t).Skip(2).First().Value).IsEqualTo("Third");
    }

    [Test]
    public async Task CanMapSymbols()
    {
        var sym1 = new XElement(W.sym, new XAttribute(W.font, "Wingdings"), new XAttribute(W._char, "F028"));
        var charFromSym1 = UnicodeMapper.SymToChar(sym1);
        var symFromChar1 = UnicodeMapper.CharToRunChild(charFromSym1);
        var sym2 = new XElement(W.sym, new XAttribute(W._char, "F028"), new XAttribute(W.font, "Wingdings"));
        var charFromSym2 = UnicodeMapper.SymToChar(sym2);
        var sym3 = new XElement(
            W.sym,
            new XAttribute(XNamespace.Xmlns + "w", W.w),
            new XAttribute(W.font, "Wingdings"),
            new XAttribute(W._char, "F028")
        );
        var charFromSym3 = UnicodeMapper.SymToChar(sym3);
        var sym4 = new XElement(
            W.sym,
            new XAttribute(XNamespace.Xmlns + "w", W.w),
            new XAttribute(W.font, "Webdings"),
            new XAttribute(W._char, "F028")
        );
        var charFromSym4 = UnicodeMapper.SymToChar(sym4);
        var symFromChar4 = UnicodeMapper.CharToRunChild(charFromSym4);
        await Assert.That(charFromSym2).IsEqualTo(charFromSym1);
        await Assert.That(charFromSym3).IsEqualTo(charFromSym1);
        await Assert.That(charFromSym4).IsNotEqualTo(charFromSym1);
        await Assert.That(symFromChar1.Attribute(W._char).Value).IsEqualTo("F028");
        await Assert.That(symFromChar1.Attribute(W.font).Value).IsEqualTo("Wingdings");
        await Assert.That(symFromChar4.Attribute(W._char).Value).IsEqualTo("F028");
        await Assert.That(symFromChar4.Attribute(W.font).Value).IsEqualTo("Webdings");
    }

    [Test]
    public async Task CanStringifySymbols()
    {
        var charFromSym1 = UnicodeMapper.SymToChar("Wingdings", '\uF028');
        var charFromSym2 = UnicodeMapper.SymToChar("Wingdings", 0xF028);
        var charFromSym3 = UnicodeMapper.SymToChar("Wingdings", "F028");
        var symFromChar1 = UnicodeMapper.CharToRunChild(charFromSym1);
        var symFromChar2 = UnicodeMapper.CharToRunChild(charFromSym2);
        var symFromChar3 = UnicodeMapper.CharToRunChild(charFromSym3);
        await Assert.That(charFromSym2).IsEqualTo(charFromSym1);
        await Assert.That(charFromSym3).IsEqualTo(charFromSym1);
        await Assert.That(symFromChar2.ToString(SaveOptions.None)).IsEqualTo(symFromChar1.ToString(SaveOptions.None));
        await Assert.That(symFromChar3.ToString(SaveOptions.None)).IsEqualTo(symFromChar1.ToString(SaveOptions.None));
    }
}
