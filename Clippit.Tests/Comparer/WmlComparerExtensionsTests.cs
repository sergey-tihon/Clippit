// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Comparer;

/// <summary>
/// Unit tests for <see cref="WmlComparerExtensions"/> covering <c>GetMainDocumentBody</c>
/// and <c>GetMainDocumentRoot</c>, including their failure modes when the main document
/// part is missing.
/// </summary>
public class WmlComparerExtensionsTests : TestsBase
{
    [Test]
    public async Task WE001_GetMainDocumentRoot_ReturnsDocumentElement()
    {
        using var stream = new MemoryStream();
        CreateEmptyWordprocessingDocument(stream);
        using var wordDocument = WordprocessingDocument.Open(stream, false);

        var root = wordDocument.GetMainDocumentRoot();

        await Assert.That(root.Name).IsEqualTo(W.document);
    }

    [Test]
    public async Task WE002_GetMainDocumentBody_ReturnsBodyElement()
    {
        using var stream = new MemoryStream();
        CreateEmptyWordprocessingDocument(stream);
        using var wordDocument = WordprocessingDocument.Open(stream, false);

        var body = wordDocument.GetMainDocumentBody();

        await Assert.That(body.Name).IsEqualTo(W.body);
    }

    [Test]
    public async Task WE003_GetMainDocumentRoot_NoMainDocumentPart_Throws()
    {
        using var stream = new MemoryStream();
        using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            // Intentionally do not add a main document part.
        }
        using var reopened = WordprocessingDocument.Open(stream, false);

        await Assert.That(() => reopened.GetMainDocumentRoot()).Throws<ArgumentException>();
    }

    [Test]
    public async Task WE004_GetMainDocumentBody_NoBodyElement_Throws()
    {
        using var stream = new MemoryStream();
        using (var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var part = wordDocument.AddMainDocumentPart();
            part.Document = new Document();
        }
        using var reopened = WordprocessingDocument.Open(stream, false);

        await Assert.That(() => reopened.GetMainDocumentBody()).Throws<ArgumentException>();
    }

    [Test]
    public async Task WE005_GetXElement_EmptyPart_Throws()
    {
        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
        var part = wordDocument.AddMainDocumentPart();

        await Assert.That(() => part.GetXElement()).Throws<ArgumentException>();
    }

    [Test]
    public async Task WE006_GetXElement_OnMainDocumentPart_ReturnsRoot()
    {
        using var stream = new MemoryStream();
        CreateEmptyWordprocessingDocument(stream);
        using var wordDocument = WordprocessingDocument.Open(stream, false);

        var root = wordDocument.MainDocumentPart.GetXElement();

        await Assert.That(root.Name).IsEqualTo(W.document);
    }
}
