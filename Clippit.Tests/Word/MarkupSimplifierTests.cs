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

    private const string CommentMarkupDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id=""1""/>
      <w:r><w:t>Hello</w:t></w:r>
      <w:commentRangeEnd w:id=""1""/>
      <w:r>
        <w:rPr><w:rStyle w:val=""CommentReference""/></w:rPr>
        <w:commentReference w:id=""1""/>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

    [Test]
    public async Task MS004_RemoveComments_RemovesInlineCommentMarkup()
    {
        var partDocument = XDocument.Parse(CommentMarkupDocumentXmlString);
        await Assert.That(partDocument.Descendants(W.commentRangeStart)).IsNotEmpty();
        await Assert.That(partDocument.Descendants(W.commentRangeEnd)).IsNotEmpty();
        await Assert.That(partDocument.Descendants(W.commentReference)).IsNotEmpty();

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { RemoveComments = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();

        await Assert.That(partDocument.Descendants(W.commentRangeStart)).IsEmpty();
        await Assert.That(partDocument.Descendants(W.commentRangeEnd)).IsEmpty();
        await Assert.That(partDocument.Descendants(W.commentReference)).IsEmpty();
        // The text run content must be preserved.
        await Assert.That(string.Concat(partDocument.Descendants(W.t).Select(t => (string)t))).IsEqualTo("Hello");
    }

    private const string BookmarksDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:bookmarkStart w:id=""1"" w:name=""MyBookmark""/>
      <w:r><w:t>Bookmarked text</w:t></w:r>
      <w:bookmarkEnd w:id=""1""/>
    </w:p>
  </w:body>
</w:document>";

    [Test]
    public async Task MS005_RemoveBookmarks_RemovesBookmarkElements()
    {
        var partDocument = XDocument.Parse(BookmarksDocumentXmlString);
        await Assert.That(partDocument.Descendants(W.bookmarkStart)).IsNotEmpty();
        await Assert.That(partDocument.Descendants(W.bookmarkEnd)).IsNotEmpty();

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { RemoveBookmarks = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();

        await Assert.That(partDocument.Descendants(W.bookmarkStart)).IsEmpty();
        await Assert.That(partDocument.Descendants(W.bookmarkEnd)).IsEmpty();
        // Text content must be preserved.
        await Assert
            .That(string.Concat(partDocument.Descendants(W.t).Select(t => (string)t)))
            .IsEqualTo("Bookmarked text");
    }

    private const string ProofErrorDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:proofErr w:type=""spellStart""/>
      <w:r><w:t>mispelled</w:t></w:r>
      <w:proofErr w:type=""spellEnd""/>
    </w:p>
  </w:body>
</w:document>";

    [Test]
    public async Task MS006_RemoveProof_RemovesProofErrElements()
    {
        var partDocument = XDocument.Parse(ProofErrorDocumentXmlString);
        await Assert.That(partDocument.Descendants(W.proofErr)).IsNotEmpty();

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { RemoveProof = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();

        await Assert.That(partDocument.Descendants(W.proofErr)).IsEmpty();
        // Text content must be preserved.
        await Assert.That(string.Concat(partDocument.Descendants(W.t).Select(t => (string)t))).IsEqualTo("mispelled");
    }

    private const string RsidDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p w:rsidR=""001234AB"" w:rsidRDefault=""001234AB"">
      <w:r w:rsidRPr=""00AB1234"">
        <w:t>Hello</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

    [Test]
    public async Task MS007_RemoveRsidInfo_RemovesRsidAttributes()
    {
        var partDocument = XDocument.Parse(RsidDocumentXmlString);
        await Assert.That(partDocument.Descendants().SelectMany(e => e.Attributes(W.rsidR))).IsNotEmpty();

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { RemoveRsidInfo = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();

        // All w:rsid* attributes must be gone.
        var rsidAttributes = partDocument
            .Descendants()
            .SelectMany(e => e.Attributes())
            .Where(a => a.Name == W.rsidR || a.Name == W.rsidDel || a.Name == W.rsidRPr || a.Name == W.rsidRDefault)
            .ToList();
        await Assert.That(rsidAttributes).IsEmpty();
        // Text content must be preserved.
        await Assert.That(string.Concat(partDocument.Descendants(W.t).Select(t => (string)t))).IsEqualTo("Hello");
    }

    private const string TabDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
  <w:body>
    <w:p>
      <w:r>
        <w:t xml:space=""preserve"">Hello</w:t>
        <w:tab/>
        <w:t>World</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>";

    [Test]
    public async Task MS008_ReplaceTabsWithSpaces_ReplacesTabElement()
    {
        var partDocument = XDocument.Parse(TabDocumentXmlString);
        await Assert.That(partDocument.Descendants(W.tab)).IsNotEmpty();

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { ReplaceTabsWithSpaces = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();

        // The w:tab element must be gone.
        await Assert.That(partDocument.Descendants(W.tab)).IsEmpty();
        // The combined text must include a space where the tab was.
        var text = string.Concat(partDocument.Descendants(W.t).Select(t => (string)t));
        await Assert.That(text).Contains(" ");
    }

    private const string HyperlinkDocumentXmlString =
        @"<w:document xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""
                     xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
  <w:body>
    <w:p>
      <w:hyperlink r:id=""rId1"" w:history=""1"">
        <w:r><w:t>Click here</w:t></w:r>
      </w:hyperlink>
    </w:p>
  </w:body>
</w:document>";

    [Test]
    public async Task MS009_RemoveHyperlinks_UnwrapsHyperlinkContent()
    {
        var partDocument = XDocument.Parse(HyperlinkDocumentXmlString);
        await Assert.That(partDocument.Descendants(W.hyperlink)).IsNotEmpty();

        using var stream = new MemoryStream();
        using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
        var part = wordDocument.AddMainDocumentPart();
        part.PutXDocument(partDocument);
        var settings = new SimplifyMarkupSettings { RemoveHyperlinks = true };
        MarkupSimplifier.SimplifyMarkup(wordDocument, settings);
        partDocument = part.GetXDocument();

        // The w:hyperlink wrapper must be gone.
        await Assert.That(partDocument.Descendants(W.hyperlink)).IsEmpty();
        // But the text content must be preserved (run is promoted to parent paragraph).
        await Assert.That(string.Concat(partDocument.Descendants(W.t).Select(t => (string)t))).IsEqualTo("Click here");
    }
}
