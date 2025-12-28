// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Xml.Linq;
using Clippit.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Word;

public class StronglyTypedBlockTests : TestsBase
{
    [Test]
    public async Task CanUseStronglyTypedBlockToDemarcateApis()
    {
        using var stream = new MemoryStream();
        CreateEmptyWordprocessingDocument(stream);
        using var wordDocument = WordprocessingDocument.Open(stream, true);
        var part = wordDocument.MainDocumentPart;
        // Add a paragraph through the PowerTools.
        var content = part.GetXDocument();
        var bodyElement = content.Descendants(W.body).First();
        bodyElement.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Added through PowerTools"))));
        part.PutXDocument();
        // This demonstrates the use of the StronglyTypedBlock in a using statement to
        // demarcate the intermittent use of the strongly typed classes.
        using (new StronglyTypedBlock(wordDocument))
        {
            // Assert that we can see the paragraph added through the PowerTools.
            var body = part.Document.Body;
            var paragraphs = body.Elements<Paragraph>().ToList();
            await Assert.That(paragraphs).HasSingleItem();
            await Assert.That(paragraphs[0].InnerText).IsEqualTo("Added through PowerTools");
            // Add a paragraph through the SDK.
            body.AppendChild(new Paragraph(new Run(new Text("Added through SDK"))));
        }

        // Assert that we can see the paragraphs added through the PowerTools and the SDK.
        content = part.GetXDocument();
        var paragraphElements = content.Descendants(W.p).ToList();
        await Assert.That(paragraphElements).HasCount(2);
        await Assert.That(paragraphElements[0].Value).IsEqualTo("Added through PowerTools");
        await Assert.That(paragraphElements[1].Value).IsEqualTo("Added through SDK");
    }

    [Test]
    public async Task ConstructorThrowsWhenPassingNull()
    {
        await Assert
            .That(() =>
            {
                using var _ = new StronglyTypedBlock(null);
            })
            .Throws<ArgumentNullException>();
    }
}
