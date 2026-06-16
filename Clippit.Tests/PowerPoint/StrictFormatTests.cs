// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using System.Text;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

public class StrictFormatTests : Clippit.Tests.TestsBase
{
    // ISO/IEC 29500 Strict uses purl.oclc.org namespaces; Transitional uses schemas.openxmlformats.org.
    private const string StrictNamespaceMarker = "http://purl.oclc.org/ooxml/";

    private static FileInfo StrictSource =>
        new(Path.Combine(new DirectoryInfo("../../../../TestFiles/").FullName, "StrictPresentation.pptx"));

    [Test]
    public async Task Fixture_is_saved_in_strict_format()
    {
        // Guards the fixture itself so the conversion tests below remain meaningful.
        var strictParts = GetXmlPartsContaining(File.ReadAllBytes(StrictSource.FullName), StrictNamespaceMarker);
        await Assert.That(strictParts).IsNotEmpty();
    }

    [Test]
    public async Task ConvertToTransitional_converts_every_part_including_nested_slide_parts()
    {
        var converted = new PmlDocument(StrictSource.FullName, convertToTransitional: true);

        var remainingStrictParts = GetXmlPartsContaining(converted.DocumentByteArray, StrictNamespaceMarker);

        await Assert.That(remainingStrictParts).IsEmpty();
    }

    [Test]
    public async Task PublishSlides_auto_converts_strict_source_to_valid_transitional_slides()
    {
        var slides = PresentationBuilder.PublishSlides(new PmlDocument(StrictSource.FullName)).ToList();

        await Assert.That(slides).IsNotEmpty();
        foreach (var slide in slides)
        {
            // No strict namespaces must survive into the published slide.
            await Assert.That(GetXmlPartsContaining(slide.DocumentByteArray, StrictNamespaceMarker)).IsEmpty();

            // And the slide must load in the typed DOM - materializing the layout is exactly
            // what throws on an unconverted / hybrid strict deck.
            using var ms = new MemoryStream();
            slide.WriteByteArray(ms);
            ms.Position = 0;
            using var doc = PresentationDocument.Open(ms, false);
            foreach (var slidePart in doc.PresentationPart!.SlideParts)
            {
                await Assert.That(slidePart.SlideLayoutPart!.SlideLayout).IsNotNull();
            }
        }
    }

    private static List<string> GetXmlPartsContaining(byte[] pptx, string marker)
    {
        using var ms = new MemoryStream(pptx);
        using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
        return zip
            .Entries.Where(e => e.FullName.EndsWith(".xml") || e.FullName.EndsWith(".rels"))
            .Where(e =>
            {
                using var reader = new StreamReader(e.Open(), Encoding.UTF8);
                return reader.ReadToEnd().Contains(marker);
            })
            .Select(e => e.FullName)
            .ToList();
    }
}
