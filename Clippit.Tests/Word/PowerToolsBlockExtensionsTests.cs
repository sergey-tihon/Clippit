// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml.Linq;
using Clippit.Core;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Word;

public class PowerToolsBlockExtensionsTests : TestsBase
{
    [Test]
    public async Task MustBeginPowerToolsBlockToUsePowerTools()
    {
        using var stream = new MemoryStream();
        CreateEmptyWordprocessingDocument(stream);

        using var wordDocument = WordprocessingDocument.Open(stream, true);
        var part = wordDocument.MainDocumentPart;

        // Add a first paragraph through the SDK.
        var body = part.Document.Body;
        body.AppendChild(new Paragraph(new Run(new Text("First"))));

        // This demonstrates the usage of the BeginPowerToolsBlock method to
        // demarcate blocks or regions of code where the PowerTools are used
        // in between usages of the strongly typed classes.
        wordDocument.BeginPowerToolsBlock();

        // Get content through the PowerTools. We will see the one paragraph added
        // by using the strongly typed SDK classes.
        var content = part.GetXDocument();
        var paragraphElements = content.Descendants(W.p).ToList();
        await Assert.That(paragraphElements).HasSingleItem();
        await Assert.That(paragraphElements[0].Value).IsEqualTo("First");

        // This demonstrates the usage of the EndPowerToolsBlock method to
        // demarcate blocks or regions of code where the PowerTools are used
        // in between usages of the strongly typed classes.
        wordDocument.EndPowerToolsBlock();

        // Add a second paragraph through the SDK in the exact same way as above.
        body = part.Document.Body;
        body.AppendChild(new Paragraph(new Run(new Text("Second"))));
        part.Document.Save();

        // Get content through the PowerTools in the exact same way as above,
        // noting that we have not used the BeginPowerToolsBlock method to
        // mark the beginning of the next PowerTools Block.
        // What we will see in this case is that we still only get the first
        // paragraph. This is caused by the GetXDocument method using the cached
        // XDocument, i.e., the annotation, rather reading the part's stream again.
        content = part.GetXDocument();
        paragraphElements = content.Descendants(W.p).ToList();
        await Assert.That(paragraphElements).HasSingleItem();
        await Assert.That(paragraphElements[0].Value).IsEqualTo("First");

        // To make the GetXDocument read the parts' streams, we need to begin
        // the next PowerTools Block. This will remove the annotations from the
        // parts and make the PowerTools read the part's stream instead of
        // using the outdated annotation.
        wordDocument.BeginPowerToolsBlock();

        // Get content through the PowerTools in the exact same way as above.
        // We should now see both paragraphs.
        content = part.GetXDocument();
        paragraphElements = content.Descendants(W.p).ToList();
        await Assert.That(paragraphElements).HasCount(2);
        await Assert.That(paragraphElements[0].Value).IsEqualTo("First");
        await Assert.That(paragraphElements[1].Value).IsEqualTo("Second");
    }

    [Test, Skip("Since v3.1 OpenXML SDK unload part root element on content stream Dispose.")]
    public async Task MustEndPowerToolsBlockToUseStronglyTypedClasses()
    {
        using var stream = new MemoryStream();
        CreateEmptyWordprocessingDocument(stream);

        using var wordDocument = WordprocessingDocument.Open(stream, true);
        var part = wordDocument.MainDocumentPart;

        // Add a paragraph through the SDK.
        var body = part.Document.Body;
        body.AppendChild(new Paragraph(new Run(new Text("Added through SDK"))));

        // Begin the PowerTools Block, which saves any changes made through the strongly
        // typed SDK classes to the parts of the WordprocessingDocument.
        // In this case, this could also be done by invoking the Save method on the
        // WordprocessingDocument, which will save all parts that had changes, or by
        // invoking part.RootElement.Save() for the one part that was changed.
        wordDocument.BeginPowerToolsBlock();

        // Add a paragraph through the PowerTools.
        var content = part.GetXDocument();
        var bodyElement = content.Descendants(W.body).First();
        bodyElement.Add(new XElement(W.p, new XElement(W.r, new XElement(W.t, "Added through PowerTools"))));
        part.PutXDocument();

        // Get the part's content through the SDK. However, we will only see what we
        // added through the SDK, not what we added through the PowerTools functionality.
        body = part.Document.Body;
        var paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs).HasSingleItem();
        await Assert.That(paragraphs[0].InnerText).IsEqualTo("Added through SDK");

        // Now, let's end the PowerTools Block, which reloads the root element of this
        // one part. Reloading those root elements this way is fine if you know exactly
        // which parts had their content changed by the Open XML PowerTools.
        wordDocument.EndPowerToolsBlock();

        // Get the part's content through the SDK. Having reloaded the root element,
        // we should now see both paragraphs.
        body = part.Document.Body;
        paragraphs = body.Elements<Paragraph>().ToList();
        await Assert.That(paragraphs).HasCount(2);
        await Assert.That(paragraphs[0].InnerText).IsEqualTo("Added through SDK");
        await Assert.That(paragraphs[1].InnerText).IsEqualTo("Added through PowerTools");
    }
}
