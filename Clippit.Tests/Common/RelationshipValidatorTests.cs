// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using System.Xml.Linq;
using Clippit.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Common;

public class RelationshipValidatorTests : TestsBase
{
    private static readonly DirectoryInfo s_testFiles = new("../../../../TestFiles");

    // RV001: a clean DOCX should produce no errors
    [Test]
    public async Task RV001_CleanDocx_NoErrors()
    {
        using var doc = WordprocessingDocument.Open(Path.Combine(s_testFiles.FullName, "Blank-wml.docx"), false);

        var errors = RelationshipValidator.Validate(doc).ToList();

        await Assert.That(errors).IsEmpty();
        await Assert.That(RelationshipValidator.IsValid(doc)).IsTrue();
    }

    // RV002: a clean PPTX should produce no errors
    [Test]
    public async Task RV002_CleanPptx_NoErrors()
    {
        using var pres = PresentationDocument.Open(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"), false);

        var errors = RelationshipValidator.Validate(pres).ToList();

        await Assert.That(errors).IsEmpty();
        await Assert.That(RelationshipValidator.IsValid(pres)).IsTrue();
    }

    // RV003: a DOCX with a dangling r:id should be detected
    [Test]
    public async Task RV003_DanglingRelationshipId_DetectedInDocx()
    {
        using var ms = new MemoryStream();

        // Build a minimal DOCX with a drawing that references a non-existent image.
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    new Paragraph(
                        new Run(
                            new Drawing(
                                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent
                                    {
                                        Cx = 1000000L,
                                        Cy = 1000000L,
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                                    {
                                        Id = 1U,
                                        Name = "Image1",
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Graphic(
                                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                                    new DocumentFormat.OpenXml.Drawing.Blip
                                                    {
                                                        // "rId999" is not registered on this part.
                                                        Embed = "rId999",
                                                    },
                                                    new DocumentFormat.OpenXml.Drawing.Stretch()
                                                ),
                                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                                    {
                                                        Id = 0U,
                                                        Name = string.Empty,
                                                    },
                                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                                ),
                                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties()
                                            )
                                        )
                                        {
                                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                                        }
                                    )
                                )
                                {
                                    DistanceFromTop = 0U,
                                    DistanceFromBottom = 0U,
                                    DistanceFromLeft = 0U,
                                    DistanceFromRight = 0U,
                                }
                            )
                        )
                    )
                )
            );
            mainPart.Document.Save();
        }

        ms.Position = 0;
        using var docRead = WordprocessingDocument.Open(ms, false);

        var errors = RelationshipValidator.Validate(docRead).ToList();

        await Assert.That(errors).Count().IsGreaterThan(0);
        await Assert.That(errors[0].RelationshipId).IsEqualTo("rId999");
        await Assert.That(RelationshipValidator.IsValid(docRead)).IsFalse();
    }

    // RV004: ArgumentNullException is thrown for null input
    [Test]
    public async Task RV004_NullPackage_ThrowsArgumentNullException()
    {
        await Assert.That(() => RelationshipValidator.Validate(null!).ToList()).Throws<ArgumentNullException>();
    }

    // RV005: a PPTX with a dangling r:id in a slide is detected
    [Test]
    public async Task RV005_DanglingRelationshipId_DetectedInPptx()
    {
        // Inject a dangling r:id into a copy of a test PPTX.
        var sourceBytes = File.ReadAllBytes(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));

        using var ms = new MemoryStream();
        ms.Write(sourceBytes);
        ms.Position = 0;

        InjectDanglingOleObjRelId(ms, "rId_dangling_test");
        ms.Position = 0;

        using var pres = PresentationDocument.Open(ms, false);

        var errors = RelationshipValidator.Validate(pres).ToList();

        // We expect at least one error for the injected dangling ID.
        await Assert.That(errors.Any(e => e.RelationshipId == "rId_dangling_test")).IsTrue();
        await Assert.That(RelationshipValidator.IsValid(pres)).IsFalse();
    }

    /// <summary>
    /// Inserts a <c>p:oleObj r:id="<paramref name="danglingId"/>"</c> element into the
    /// first slide's relationship XML to simulate a dangling relationship reference.
    /// </summary>
    private static void InjectDanglingOleObjRelId(Stream pptxStream, string danglingId)
    {
        XNamespace pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var zip = new ZipArchive(pptxStream, ZipArchiveMode.Update, leaveOpen: true);

        // Find the first slide entry.
        var slideEntry = zip.Entries.FirstOrDefault(e =>
            e.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
            && e.FullName.EndsWith(".xml", StringComparison.Ordinal)
        );

        if (slideEntry is null)
            return;

        XDocument xDoc;
        using (var s = slideEntry.Open())
            xDoc = XDocument.Load(s);

        // Append a <p:oleObj r:id="danglingId"> to the <p:spTree> if it exists,
        // otherwise to the document root.
        var spTree = xDoc.Descendants(pNs + "spTree").FirstOrDefault();
        var target = spTree ?? xDoc.Root;
        target?.Add(new XElement(pNs + "oleObj", new XAttribute(rNs + "id", danglingId)));

        // Replace the entry with the modified XML.
        var fullName = slideEntry.FullName;
        slideEntry.Delete();
        var newEntry = zip.CreateEntry(fullName);
        using var writer = new System.IO.StreamWriter(newEntry.Open());
        using var xmlWriter = System.Xml.XmlWriter.Create(writer);
        xDoc.WriteTo(xmlWriter);
    }
}
