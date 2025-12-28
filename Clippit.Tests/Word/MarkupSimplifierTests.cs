// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Xml.Linq;
using Clippit.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Word;

public class MarkupSimplifierTests
{
    private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;
    private const string SmartTagDocumentTextValue = "The countries include Algeria, Botswana, and Sri Lanka.";
    private const string SmartTagDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p >
      <w:r>
        <w:t xml:space=""preserve"">The countries include </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:r>
          <w:t>Algeria</w:t>
        </w:r>
      </w:smartTag>
      <w:r>
        <w:t xml:space=""preserve"">, </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:r>
          <w:t>Botswana</w:t>
        </w:r>
      </w:smartTag>
      <w:r>
        <w:t xml:space=""preserve"">, and </w:t>
      </w:r>
      <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""country-region"">
        <w:smartTag w:uri=""urn:schemas-microsoft-com:office:smarttags"" w:element=""place"">
          <w:r>
            <w:t>Sri Lanka</w:t>
          </w:r>
        </w:smartTag>
      </w:smartTag>
      <w:r>
        <w:t>.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
";
    private const string SdtDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:sdt>
      <w:sdtPr>
        <w:text/>
      </w:sdtPr>
      <w:sdtContent>
        <w:p>
          <w:r>
            <w:t>Hello World!</w:t>
          </w:r>
        </w:p>
      </w:sdtContent>
    </w:sdt>
  </w:body>
</w:document>";
    private const string GoBackBookmarkDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id=""0"" w:name=""_GoBack""/>
      <w:bookmarkEnd w:id=""0""/>
    </w:p>
  </w:body>
</w:document>";

    [Test]
    public async Task CanRemoveSmartTags()
    {
        var partDocument = XDocument.Parse(SmartTagDocumentXmlString);
        await Assert.That(partDocument.Descendants(W.smartTag)).IsNotEmpty();

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { RemoveSmartTags = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();
        var t = partDocument.Descendants(W.t).First();
        await Assert.That(partDocument.Descendants(W.smartTag)).IsEmpty();
        await Assert.That(t.Value).IsEqualTo(SmartTagDocumentTextValue);
    }

    [Test]
    public async Task CanRemoveContentControls()
    {
        var partDocument = XDocument.Parse(SdtDocumentXmlString);
        await Assert.That(partDocument.Descendants(W.sdt)).IsNotEmpty();

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { RemoveContentControls = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();
        var element = partDocument.Descendants(W.body).Descendants().First();
        await Assert.That(partDocument.Descendants(W.sdt)).IsEmpty();
        await Assert.That(element.Name).IsEqualTo(W.p);
    }

    [Test]
    public async Task CanRemoveGoBackBookmarks()
    {
        var partDocument = XDocument.Parse(GoBackBookmarkDocumentXmlString);
        await Assert
            .That(partDocument.Descendants(W.bookmarkStart))
            .Contains(e => e.Attribute(W.name).Value == "_GoBack" && e.Attribute(W.id).Value == "0");
        await Assert.That(partDocument.Descendants(W.bookmarkEnd)).Contains(e => e.Attribute(W.id).Value == "0");
        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { RemoveGoBackBookmark = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();
        await Assert.That(partDocument.Descendants(W.bookmarkStart)).IsEmpty();
        await Assert.That(partDocument.Descendants(W.bookmarkEnd)).IsEmpty();
    }
}
