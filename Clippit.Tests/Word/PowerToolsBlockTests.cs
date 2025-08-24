// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Xml.Linq;
using Clippit.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Word;

public class PowerToolsBlockTests : TestsBase
{
    [Test]
    public async Task CanUsePowerToolsBlockToDemarcateApis()
    {
        using var stream = new MemoryStream();
        CreateEmptyWordprocessingDocument(stream);
        using var wordDocument = WordprocessingDocument.Open(stream, true);
        var part = wordDocument.MainDocumentPart;
        // Add a paragraph through the SDK.
        var body = part.Document.Body;
        body.AppendChild(new Paragraph(new Run(new Text("Added through SDK"))));
        // This demonstrates the use of the PowerToolsBlock in a using statement to
        // demarcate the intermittent use of the PowerTools.
        using (new PowerToolsBlock(wordDocument))
        {
            // Assert that we can see the paragraph added through the strongly typed classes.
            var content = part.GetXDocument();
            var paragraphElements = content.Descendants(W.p).ToList();
            await Assert.That(paragraphElements).HasSingleItem();
            await Assert.That(paragraphElements[0].Value).IsEqualTo("Added through SDK");
            // Add a paragraph through the PowerTools.
            var bodyElement = content.Descendants(W.body).First();
            bodyElement.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Added through PowerTools"))));
            part.PutXDocument();
        }

        // Get the part's content through the SDK. Having used the PowerToolsBlock,
        // we should see both paragraphs.
        body = part.Document.Body;
        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs).HasCount(2);
        await Assert.That(paragraphs[0].InnerText).IsEqualTo("Added through SDK");
        await Assert.That(paragraphs[0].InnerText).IsEqualTo("Added through PowerTools");
    }

    [Test]
    public async Task ConstructorThrowsWhenPassingNull()
    {
        await Assert
            .That(() =>
            {
                using var _ = new PowerToolsBlock(null);
            })
            .Throws<ArgumentNullException>();
    }
}
