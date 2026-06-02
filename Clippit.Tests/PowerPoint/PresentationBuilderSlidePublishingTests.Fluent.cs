using System.IO.Compression;
using System.Xml.Linq;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

public partial class PresentationBuilderSlidePublishingTests
{
    /// <summary>
    /// Regression test for https://github.com/sergey-tihon/Clippit/issues/233 —
    /// AddSlidePart must not throw InvalidDataException when a slide contains an image part
    /// whose ZIP entry has a corrupt local file header.
    /// </summary>
    [Test]
    public async Task AddSlidePart_WithCorruptImageLocalFileHeader_DoesNotThrow()
    {
        var sourcePath = Path.Combine(SourceDirectory, "BRK3066.pptx");
        var openSettings = new OpenSettings { AutoSave = false };

        // Copy the source file into a writable memory stream, then corrupt an image entry's
        // local file header to reproduce the InvalidDataException reported in issue #233.
        using var srcMemory = new MemoryStream();
        await using (var fs = File.OpenRead(sourcePath))
            await fs.CopyToAsync(srcMemory);

        var corrupted = CorruptPptMediaLocalFileHeader(srcMemory.ToArray());
        using var corruptedMemory = new MemoryStream(corrupted);

        using var srcDoc = PresentationDocument.Open(corruptedMemory, false, openSettings);
        ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);

        var firstSlideId = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc).First();
        var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(firstSlideId);

        // Should not throw InvalidDataException
        using var destStream = new MemoryStream();
        using (var destDoc = PresentationBuilder.NewDocument(destStream))
        using (var builder = PresentationBuilder.Create(destDoc))
        {
            builder.AddSlidePart(srcSlidePart);
        }

        await Assert.That(destStream.Length).IsGreaterThan(0);
    }

    /// <summary>
    /// Scans raw ZIP bytes for the first entry whose path starts with "ppt/media/"
    /// and corrupts its local file header signature so that <see cref="ZipArchiveEntry.Open"/>
    /// throws <see cref="InvalidDataException"/> when the entry is read.
    /// </summary>
    private static byte[] CorruptPptMediaLocalFileHeader(byte[] zipBytes)
    {
        // ZIP local file header layout:
        //   [0-3]  signature  0x04034B50
        //   [26-27] file name length (little-endian uint16)
        //   [28-29] extra field length (little-endian uint16)
        //   [30+]   file name
        const uint LocalFileHeaderSignature = 0x04034B50;

        for (var i = 0; i <= zipBytes.Length - 30; i++)
        {
            var sig =
                zipBytes[i]
                | ((uint)zipBytes[i + 1] << 8)
                | ((uint)zipBytes[i + 2] << 16)
                | ((uint)zipBytes[i + 3] << 24);
            if (sig != LocalFileHeaderSignature)
                continue;

            var nameLen = zipBytes[i + 26] | (zipBytes[i + 27] << 8);
            if (i + 30 + nameLen > zipBytes.Length)
                continue;

            var name = System.Text.Encoding.UTF8.GetString(zipBytes, i + 30, nameLen);
            if (!name.StartsWith("ppt/media/", StringComparison.Ordinal))
                continue;

            // Corrupt bytes 2-3 of the signature so Open() throws InvalidDataException.
            var result = (byte[])zipBytes.Clone();
            result[i + 2] = 0xFF;
            result[i + 3] = 0xFF;
            return result;
        }

        throw new InvalidOperationException("No ppt/media/ entry found in the PPTX ZIP archive.");
    }

    /// <summary>
    /// Regression test for https://github.com/sergey-tihon/Clippit/issues/155 —
    /// AddSlidePart must not throw KeyNotFoundException when a slide contains a p:oleObj
    /// or p:externalData element whose r:id is not found as either a part or an external relationship.
    /// </summary>
    [Test]
    [Arguments("oleObj")]
    [Arguments("externalData")]
    public async Task AddSlidePart_WithDanglingOleObjOrExternalDataRelationship_DoesNotThrow(string elementLocalName)
    {
        var sourcePath = Path.Combine(SourceDirectory, "BRK3066.pptx");
        var openSettings = new OpenSettings { AutoSave = false };

        // Copy the source file into a writable memory stream so we can inject a dangling reference.
        using var srcMemory = new MemoryStream();
        await using (var fs = File.OpenRead(sourcePath))
            await fs.CopyToAsync(srcMemory);
        srcMemory.Position = 0;

        using var srcDoc = PresentationDocument.Open(srcMemory, true, openSettings);
        ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);

        var firstSlideId = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc).First();
        var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(firstSlideId);

        // Inject the element with a relationship ID that doesn't exist as a part or external relationship.
        // This simulates a slide produced by third-party software with a dangling reference.
        XNamespace pns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var slideXDoc = srcSlidePart.GetXDocument();
        var spTree = slideXDoc.Descendants(pns + "spTree").FirstOrDefault();
        ArgumentNullException.ThrowIfNull(spTree);
        spTree.Add(new XElement(pns + elementLocalName, new XAttribute(rns + "id", "rId_dangling_999")));
        srcSlidePart.PutXDocument(slideXDoc);

        // Should not throw KeyNotFoundException
        using var destStream = new MemoryStream();
        using (var destDoc = PresentationBuilder.NewDocument(destStream))
        using (var builder = PresentationBuilder.Create(destDoc))
        {
            builder.AddSlidePart(srcSlidePart);
        }

        await Assert.That(destStream.Length).IsGreaterThan(0);
    }

    [Test]
    [MethodDataSource(typeof(PublishingTestData), nameof(PublishingTestData.Files))]
    public async Task PublishUsingMemDocs(string sourcePath, CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var targetDir = Path.Combine(TargetDirectory, fileName);
        if (Directory.Exists(targetDir))
            Directory.Delete(targetDir, true);
        Directory.CreateDirectory(targetDir);
        await using var srcStream = File.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var openSettings = new OpenSettings { AutoSave = false };
        using var srcDoc = OpenXmlExtensions.OpenPresentation(srcStream, false, openSettings);
        ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);
        var slideNumber = 0;
        var slidesIds = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc);
        foreach (var slideId in slidesIds)
        {
            var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slideId);
            var title = PresentationBuilderTools.GetSlideTitle(srcSlidePart.GetXElement());
            using var stream = new MemoryStream();
            using (var newDocument = PresentationBuilder.NewDocument(stream))
            {
                using (var builder = PresentationBuilder.Create(newDocument))
                {
                    var newSlidePart = builder.AddSlidePart(srcSlidePart);
                    // Remove the show attribute from the slide element (if it exists)
                    var slideDocument = newSlidePart.GetXDocument();
                    slideDocument.Root?.Attribute(NoNamespace.show)?.Remove();
                }

                // Set the title of the new presentation to the title of the slide
                newDocument.PackageProperties.Title = title;
                await ValidateRelationships(newDocument);
            }

            var slideFileName = string.Concat(fileName, $"_{++slideNumber:000}.pptx");
            await using var fs = File.Create(Path.Combine(targetDir, slideFileName));
            stream.Position = 0;
            await stream.CopyToAsync(fs, cancellationToken).ConfigureAwait(false);
            srcSlidePart.RemoveAnnotations<XDocument>();
            srcSlidePart.UnloadRootElement();
        }

        Console.WriteLine($"GC Total Memory: {GC.GetTotalMemory(false) / 1024 / 1024} MB");
    }

    /// <summary>
    /// Regression test for https://github.com/sergey-tihon/Clippit/issues/231 —
    /// AddSlidePart must not throw ArgumentException ("An invalid XML ID string")
    /// when a slide contains a p:custData element whose CustomXmlPart has a CustomXmlPropertiesPart.
    /// </summary>
    [Test]
    public async Task AddSlidePart_WithCustDataHavingCustomXmlPropertiesPart_DoesNotThrow()
    {
        var sourcePath = Path.Combine(SourceDirectory, "BRK3066.pptx");
        var openSettings = new OpenSettings { AutoSave = false };

        // Copy the source file into a writable memory stream so we can inject the custom-data parts.
        using var srcMemory = new MemoryStream();
        await using (var fs = File.OpenRead(sourcePath))
            await fs.CopyToAsync(srcMemory);
        srcMemory.Position = 0;

        using var srcDoc = PresentationDocument.Open(srcMemory, true, openSettings);
        ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);

        var firstSlideId = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc).First();
        var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(firstSlideId);

        // Attach a CustomXmlPart (with a CustomXmlPropertiesPart child) to the slide and
        // add a p:custData element referencing it.  This is the exact combination that used
        // to throw "An invalid XML ID string" inside CopyRelatedPartsForContentParts because
        // AddNewPart<CustomXmlPropertiesPart> was incorrectly called with a content-type
        // string as the relationship-ID argument.
        var customXmlPart = srcSlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        customXmlPart.PutXDocument(new XDocument(new XElement("root", "test")));

        var propsPart = customXmlPart.AddNewPart<CustomXmlPropertiesPart>();
        propsPart.PutXDocument(
            new XDocument(
                new XElement(
                    XName.Get("datastoreItem", "http://schemas.openxmlformats.org/officeDocument/2006/customXml"),
                    new XAttribute(
                        XName.Get("itemID", "http://schemas.openxmlformats.org/officeDocument/2006/customXml"),
                        "{12345678-1234-1234-1234-123456789012}"
                    )
                )
            )
        );

        // Inject the p:custData element into the slide XML.
        XNamespace pns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var custDataRelId = srcSlidePart.GetIdOfPart(customXmlPart);
        var slideXDoc = srcSlidePart.GetXDocument();
        slideXDoc.Root?.Add(
            new XElement(pns + "custDataLst", new XElement(pns + "custData", new XAttribute(rns + "id", custDataRelId)))
        );
        srcSlidePart.PutXDocument(slideXDoc);

        // Should not throw ArgumentException ("An invalid XML ID string")
        using var destStream = new MemoryStream();
        using (var destDoc = PresentationBuilder.NewDocument(destStream))
        using (var builder = PresentationBuilder.Create(destDoc))
        {
            builder.AddSlidePart(srcSlidePart);
        }

        await Assert.That(destStream.Length).IsGreaterThan(0);
    }

    /// <summary>
    /// Regression test: a slide whose p:custData element references an empty (zero-byte)
    /// CustomXmlPart should be copied without error, and the empty part must not appear
    /// in the output (it has no root element and causes a PowerPoint repair dialog).
    /// </summary>
    [Test]
    public async Task AddSlidePart_WithEmptyCustDataPart_SkipsEmptyPartAndDoesNotThrow()
    {
        var sourcePath = Path.Combine(SourceDirectory, "BRK3066.pptx");
        var openSettings = new OpenSettings { AutoSave = false };

        using var srcMemory = new MemoryStream();
        await using (var fs = File.OpenRead(sourcePath))
            await fs.CopyToAsync(srcMemory);
        srcMemory.Position = 0;

        using var srcDoc = PresentationDocument.Open(srcMemory, true, openSettings);
        ArgumentNullException.ThrowIfNull(srcDoc.PresentationPart);

        var firstSlideId = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc).First();
        var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(firstSlideId);

        // Attach an empty CustomXmlPart (zero bytes, no root element) to the slide.
        var emptyXmlPart = srcSlidePart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
        // deliberately leave the stream empty (no PutXDocument call)

        var propsPart = emptyXmlPart.AddNewPart<CustomXmlPropertiesPart>();
        propsPart.PutXDocument(
            new XDocument(
                new XElement(
                    XName.Get("datastoreItem", "http://schemas.openxmlformats.org/officeDocument/2006/customXml"),
                    new XAttribute(
                        XName.Get("itemID", "http://schemas.openxmlformats.org/officeDocument/2006/customXml"),
                        "{7D2CF446-7740-064C-89BA-4FB9111814D7}"
                    )
                )
            )
        );

        XNamespace pns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        var custDataRelId = srcSlidePart.GetIdOfPart(emptyXmlPart);
        var slideXDoc = srcSlidePart.GetXDocument();
        slideXDoc.Root?.Add(
            new XElement(pns + "custDataLst", new XElement(pns + "custData", new XAttribute(rns + "id", custDataRelId)))
        );
        srcSlidePart.PutXDocument(slideXDoc);

        using var destStream = new MemoryStream();
        // Should not throw
        using (var destDoc = PresentationBuilder.NewDocument(destStream))
        using (var builder = PresentationBuilder.Create(destDoc))
        {
            builder.AddSlidePart(srcSlidePart);
        }

        await Assert.That(destStream.Length).IsGreaterThan(0);

        // The empty customXml part must not appear in the output package.
        destStream.Position = 0;
        using var archive = new System.IO.Compression.ZipArchive(
            destStream,
            System.IO.Compression.ZipArchiveMode.Read,
            leaveOpen: true
        );
        var customXmlEntries = archive
            .Entries.Where(e =>
                e.FullName.StartsWith("customXml/item", StringComparison.Ordinal)
                && e.FullName.EndsWith(".xml", StringComparison.Ordinal)
                && !e.FullName.Contains("Props", StringComparison.Ordinal)
            )
            .ToList();
        await Assert.That(customXmlEntries).IsEmpty();

        // The <p:custDataLst> element must also be pruned from the slide XML
        // so no dangling relationship reference remains.
        destStream.Position = 0;
        using var doc = PresentationDocument.Open(destStream, false);
        var slidePart = doc.PresentationPart!.SlideParts.First();
        XNamespace pns2 = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var custDataLst = slidePart.GetXDocument().Root?.Element(pns2 + "custDataLst");
        await Assert.That(custDataLst).IsNull();
    }

    [Test]
    [MethodDataSource(typeof(PublishingTestData), nameof(PublishingTestData.Files))]
    public async Task MergeAllPowerPointBack(string sourcePath, CancellationToken cancellationToken)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var targetDir = Path.Combine(TargetDirectory, fileName);
        if (!Directory.Exists(targetDir))
        {
            Console.WriteLine("Directory not found: " + targetDir);
            return;
        }

        var slides = Directory.GetFiles(targetDir, "*.pptx", SearchOption.TopDirectoryOnly);
        if (slides.Length < 1)
        {
            Console.WriteLine("Not enough slides to merge.");
            return;
        }

        Array.Sort(slides);
        // Create a memory stream from the original presentation
        using var ms = new MemoryStream();
        await using (var fs = File.OpenRead(sourcePath))
        {
            await fs.CopyToAsync(ms, cancellationToken).ConfigureAwait(false);
        }

        // Use the first slide as the base document
        var setting = new OpenSettings { AutoSave = false };
        using (var baseDoc = PresentationDocument.Open(ms, true, setting))
        {
            ArgumentNullException.ThrowIfNull(baseDoc.PresentationPart);
            // Merge the remaining slides into the base document (one builder instance)
            using (var builder = PresentationBuilder.Create(baseDoc))
            {
                foreach (var path in slides)
                {
                    using var doc = PresentationDocument.Open(path, false, setting);
                    ArgumentNullException.ThrowIfNull(doc.PresentationPart);
                    // Add all slides in the correct order
                    foreach (var slidePath in PresentationBuilderTools.GetSlideIdsInOrder(doc))
                    {
                        var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slidePath);
                        builder.AddSlidePart(slidePart);
                    }
                }
            }

            baseDoc.PackageProperties.Title = $"{fileName} - Merged Deck X2";
        }

        // Save the merged document to a file
        var resultFile = Path.Combine(TargetDirectory, $"{fileName}_MergedDeckX2.pptx");
        ms.Position = 0;
        await using var resFile = File.Create(resultFile);
        await ms.CopyToAsync(resFile, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Regression test for https://github.com/sergey-tihon/Clippit/issues/286 —
    /// GetSlideTitle must return an empty string (not throw NullReferenceException)
    /// when the slide element is missing a p:cSld or p:spTree child.
    /// </summary>
    [Test]
    public async Task GetSlideTitle_MissingCsld_ReturnsEmptyString()
    {
        // Slide element with no p:cSld child at all
        var slide = new XElement(P.sld);
        var title = PresentationBuilderTools.GetSlideTitle(slide);
        await Assert.That(title).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task GetSlideTitle_MissingSpTree_ReturnsEmptyString()
    {
        // p:cSld present but without p:spTree
        var slide = new XElement(P.sld, new XElement(P.cSld));
        var title = PresentationBuilderTools.GetSlideTitle(slide);
        await Assert.That(title).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task GetSlideTitle_NoTitleShape_ReturnsEmptyString()
    {
        // p:cSld/p:spTree present but no shapes with title placeholder
        var slide = new XElement(P.sld, new XElement(P.cSld, new XElement(P.spTree)));
        var title = PresentationBuilderTools.GetSlideTitle(slide);
        await Assert.That(title).IsEqualTo(string.Empty);
    }

    [Test]
    public async Task GetSlideTitle_TitleShapeWithoutTextBody_ReturnsEmptyString()
    {
        var slide = new XElement(
            P.sld,
            new XElement(
                P.cSld,
                new XElement(
                    P.spTree,
                    new XElement(
                        P.sp,
                        new XElement(
                            P.nvSpPr,
                            new XElement(P.nvPr, new XElement(P.ph, new XAttribute(NoNamespace.type, "title")))
                        )
                    )
                )
            )
        );

        var title = PresentationBuilderTools.GetSlideTitle(slide);

        await Assert.That(title).IsEqualTo(string.Empty);
    }
}
