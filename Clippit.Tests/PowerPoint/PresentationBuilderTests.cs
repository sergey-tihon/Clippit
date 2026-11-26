// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using System.Xml.Linq;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

public class PresentationBuilderTests : TestsBase
{
    [Test]
    public async Task PB001_Formatting()
    {
        var name1 = "PB001-Input1.pptx";
        var name2 = "PB001-Input2.pptx";
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
        var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
        var sources = new List<SlideSource>
        {
            new(new PmlDocument(source1Pptx.FullName), 1, true),
            new(new PmlDocument(source2Pptx.FullName), 0, true),
        };
        var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB001-Formatting.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        using var pptx = PresentationDocument.Open(processedDestPptx.FullName, false);
        await Validate(pptx);
    }

    [Test]
    public async Task PB002_Formatting()
    {
        var name2 = "PB001-Input2.pptx";
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
        var sources = new List<SlideSource> { new(new PmlDocument(source2Pptx.FullName), 0, true) };
        var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB002-Formatting.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        using var pptx = PresentationDocument.Open(processedDestPptx.FullName, false);
        await Validate(pptx);
    }

    [Test]
    public async Task PB003_Formatting()
    {
        var name1 = "PB001-Input1.pptx";
        var name2 = "PB001-Input3.pptx";
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
        var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
        var sources = new List<SlideSource>
        {
            new(new PmlDocument(source1Pptx.FullName), 1, true),
            new(new PmlDocument(source2Pptx.FullName), 0, true),
        };
        var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB003-Formatting.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        using var pptx = PresentationDocument.Open(processedDestPptx.FullName, false);
        await Validate(pptx);
    }

    [Test]
    public async Task PB004_Formatting()
    {
        var name1 = "PB001-Input1.pptx";
        var name2 = "PB001-Input3.pptx";
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
        var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
        var sources = new List<SlideSource>
        {
            new(new PmlDocument(source2Pptx.FullName), 0, true),
            new(new PmlDocument(source1Pptx.FullName), 1, true),
        };
        var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB004-Formatting.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        using var pptx = PresentationDocument.Open(processedDestPptx.FullName, false);
        await Validate(pptx);
    }

    [Test]
    public async Task PB005_Formatting()
    {
        var name1 = "PB001-Input1.pptx";
        var name2 = "PB001-Input3.pptx";
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
        var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));
        var sources = new List<SlideSource>
        {
            new(new PmlDocument(source2Pptx.FullName), 0, 0, true),
            new(new PmlDocument(source1Pptx.FullName), 1, true),
            new(new PmlDocument(source2Pptx.FullName), 0, true),
        };
        var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB005-Formatting.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        using var pptx = PresentationDocument.Open(processedDestPptx.FullName, false);
        await Validate(pptx);
    }

    [Test]
    public async Task PB006_VideoFormats()
    {
        // This presentation contains videos with content types video/mp4, video/quicktime, video/unknown, video/x-ms-asf, and video/x-msvideo.
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourcePptx = new FileInfo(Path.Combine(sourceDir.FullName, "PP006-Videos.pptx"));
        var oldMediaDataContentTypes = GetMediaDataContentTypes(sourcePptx);
        var sources = new List<SlideSource> { new(new PmlDocument(sourcePptx.FullName), true) };
        var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB006-Videos.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        var newMediaDataContentTypes = GetMediaDataContentTypes(processedDestPptx);
        await Assert.That(newMediaDataContentTypes).IsEquivalentTo(oldMediaDataContentTypes);
    }

    [Test]
    public async Task PB007_ChartWithExtendedPartWorkbook()
    {
        // Regression test for #388: chart workbook stored as ExtendedPart must be copied correctly.
        // The fixture uses a non-standard relationship type that the SDK resolves as ExtendedPart.
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourcePptx = new FileInfo(Path.Combine(sourceDir.FullName, "PB007-Chart-ExtendedPart-Workbook.pptx"));
        var sources = new List<SlideSource> { new(new PmlDocument(sourcePptx.FullName), true) };
        var outputPptx = new FileInfo(Path.Combine(TempDir, "PB007-Output.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(outputPptx.FullName);

        using var pptx = PresentationDocument.Open(outputPptx.FullName, false);
        await Validate(pptx);

        var chartPart = pptx.PresentationPart!.SlideParts.Single().ChartParts.Single();
        var extendedParts = chartPart.Parts.Where(p => p.OpenXmlPart is ExtendedPart).ToList();
        await Assert.That(extendedParts).HasCount().EqualTo(1);
    }

    [Test]
    public async Task PB008_SlideWith3DModel()
    {
        // A 3D model is an <am3d:model3d r:embed> that points to a .glb part, with a raster
        // preview in <am3d:blip r:embed>. Both relationships must exist in the extracted slide.
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourcePptx = new FileInfo(Path.Combine(sourceDir.FullName, "PB008-3DModel.pptx"));
        byte[] sourceModel;
        using (var source = PresentationDocument.Open(sourcePptx.FullName, false))
        {
            var sourceSlide = source.PresentationPart!.SlideParts.Single(s =>
                s.Model3DReferenceRelationshipParts.Any()
            );
            sourceModel = ReadAllBytes(sourceSlide.Model3DReferenceRelationshipParts.Single());
        }

        var sources = new List<SlideSource> { new(new PmlDocument(sourcePptx.FullName), 1, 1, true) };
        var outputPptx = new FileInfo(Path.Combine(TempDir, "PB008-Output.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(outputPptx.FullName);

        using var pptx = PresentationDocument.Open(outputPptx.FullName, false);
        await Validate(pptx);

        var slidePart = pptx.PresentationPart!.SlideParts.Single();
        var modelPart = slidePart.Model3DReferenceRelationshipParts.Single();
        await Assert.That(ReadAllBytes(modelPart)).IsEquivalentTo(sourceModel);

        var slideXml = slidePart.GetXDocument();
        var model3d = slideXml.Descendants(AM3D.model3d).Single();
        await Assert.That((string)model3d.Attribute(R.embed)).IsEqualTo(slidePart.GetIdOfPart(modelPart));

        var rasterId = (string)slideXml.Descendants(AM3D.blip).Single().Attribute(R.embed);
        await Assert.That(slidePart.GetPartById(rasterId)).IsTypeOf<ImagePart>();
    }

    [Test]
    public async Task PB009_IgnorablePrefixStaysDeclared()
    {
        // PowerPoint for Mac writes mc:Ignorable="mv" on the presentation root. The prefix must
        // stay declared in the output, or the Ignorable attribute is invalid.
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourcePptx = new FileInfo(Path.Combine(sourceDir.FullName, "PB009-MacVmlIgnorable.pptx"));
        var sources = new List<SlideSource> { new(new PmlDocument(sourcePptx.FullName), 1, 1, true) };
        var outputPptx = new FileInfo(Path.Combine(TempDir, "PB009-Output.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(outputPptx.FullName);

        using var pptx = PresentationDocument.Open(outputPptx.FullName, false);
        await Validate(pptx);

        var root = pptx.PresentationPart!.GetXDocument().Root!;
        var ignorable = ((string)root.Attribute(MC.Ignorable)).Split(' ', StringSplitOptions.RemoveEmptyEntries);
        foreach (var prefix in ignorable)
            await Assert.That(root.GetNamespaceOfPrefix(prefix)).IsNotNull();
    }

    [Test]
    public async Task PB010_EmbeddedFontWithEot10Header()
    {
        var eot10 = EmbeddedFontHeaderTests.CreateEot(0x00010000, EmbeddedFontHeaderTests.FontData);
        var eot22 = EmbeddedFontHeaderTests.CreateEot(0x00020002, EmbeddedFontHeaderTests.FontData, new byte[24]);

        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        using var sourceStream = new MemoryStream();
        sourceStream.Write(File.ReadAllBytes(Path.Combine(sourceDir.FullName, "PB001-Input1.pptx")));
        using (var source = PresentationDocument.Open(sourceStream, true))
        {
            var presentationPart = source.PresentationPart!;
            var fontList = new XElement(P.embeddedFontLst);
            foreach (
                var (typeface, style, data) in new[] { ("Old Font", P.regular, eot10), ("New Font", P.bold, eot22) }
            )
            {
                var fontPart = presentationPart.AddFontPart(FontPartType.FontData);
                using (var fontStream = new MemoryStream(data))
                    fontPart.FeedData(fontStream);
                fontList.Add(
                    new XElement(
                        P.embeddedFont,
                        new XElement(P.font, new XAttribute("typeface", typeface)),
                        new XElement(style, new XAttribute(R.id, presentationPart.GetIdOfPart(fontPart)))
                    )
                );
            }
            var root = presentationPart.GetXDocument().Root!;
            root.Element(P.notesSz)!.AddAfterSelf(fontList);
            presentationPart.PutXDocument();
        }

        var sources = new List<SlideSource>
        {
            new(new PmlDocument("PB010-Input.pptx", sourceStream.ToArray()), 0, 1, true),
        };
        var outputPptx = new FileInfo(Path.Combine(TempDir, "PB010-Output.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(outputPptx.FullName);

        using var pptx = PresentationDocument.Open(outputPptx.FullName, false);
        await Validate(pptx);

        var fonts = pptx.PresentationPart!.GetXDocument()
            .Root!.Element(P.embeddedFontLst)!
            .Elements(P.embeddedFont)
            .ToDictionary(
                f => (string)f.Element(P.font)!.Attribute("typeface"),
                f => ReadAllBytes(pptx.PresentationPart.GetPartById((string)f.Elements().Last().Attribute(R.id)))
            );

        await Assert.That(fonts["Old Font"]).IsEquivalentTo(EmbeddedFontHeader.UpgradeToVersion22(eot10)!);
        await Assert.That(fonts["New Font"]).IsEquivalentTo(eot22);
    }

    [Test]
    public async Task PB011_SlideWith3DModelSourceUrl()
    {
        var sourceUrl = new Uri("https://example.com/models/cube");

        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        using var sourceStream = new MemoryStream();
        sourceStream.Write(File.ReadAllBytes(Path.Combine(sourceDir.FullName, "PB008-3DModel.pptx")));
        using (var source = PresentationDocument.Open(sourceStream, true))
        {
            var sourceSlide = source.PresentationPart!.SlideParts.Single(s =>
                s.Model3DReferenceRelationshipParts.Any()
            );
            var hyperlink = sourceSlide.AddHyperlinkRelationship(sourceUrl, true);
            var raster = sourceSlide.GetXDocument().Descendants(AM3D.blip).Single().Parent!;
            raster.AddBeforeSelf(new XElement(AM3D.attrSrcUrl, new XAttribute(R.id, hyperlink.Id)));
            sourceSlide.PutXDocument();
        }

        var sources = new List<SlideSource>
        {
            new(new PmlDocument("PB011-Input.pptx", sourceStream.ToArray()), 1, 1, true),
        };
        var outputPptx = new FileInfo(Path.Combine(TempDir, "PB011-Output.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(outputPptx.FullName);

        using var pptx = PresentationDocument.Open(outputPptx.FullName, false);
        await Validate(pptx);

        var slidePart = pptx.PresentationPart!.SlideParts.Single();
        var relId = (string)slidePart.GetXDocument().Descendants(AM3D.attrSrcUrl).Single().Attribute(R.id);
        var copied = slidePart.HyperlinkRelationships.Single(h => h.Id == relId);
        await Assert.That(copied.Uri).IsEqualTo(sourceUrl);
        await Assert.That(copied.IsExternal).IsTrue();
    }

    [Test]
    public async Task PB012_SlideWithWebExtension()
    {
        var snapshotUrl = new Uri("https://example.com/snapshot.png");
        var snapshot = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg=="
        );
        var webExtension = new XElement(
            WE.webextension,
            new XAttribute(XNamespace.Xmlns + "we", WE.we),
            new XAttribute(XNamespace.Xmlns + "r", R.r),
            new XAttribute("id", "{6F1B2A5C-0B5E-4C0E-9A57-2D0C3E6B7A11}"),
            new XElement(
                WE.reference,
                new XAttribute("id", "wa000000000"),
                new XAttribute("version", "1.0.0.0"),
                new XAttribute("store", "en-US"),
                new XAttribute("storeType", "OMEX")
            ),
            new XElement(WE.alternateReferences),
            new XElement(WE.properties),
            new XElement(WE.bindings),
            new XElement(WE.snapshot)
        );

        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        using var sourceStream = new MemoryStream();
        sourceStream.Write(File.ReadAllBytes(Path.Combine(sourceDir.FullName, "PB001-Input1.pptx")));
        using (var source = PresentationDocument.Open(sourceStream, true))
        {
            var presentationPart = source.PresentationPart!;
            var firstSlideId = (string)
                presentationPart.GetXDocument().Root!.Element(P.sldIdLst)!.Elements(P.sldId).First().Attribute(R.id);
            var sourceSlide = (SlidePart)presentationPart.GetPartById(firstSlideId);

            var webExtensionPart = sourceSlide.AddNewPart<WebExtensionPart>();
            var snapshotPart = webExtensionPart.AddImagePart(ImagePartType.Png);
            using (var snapshotStream = new MemoryStream(snapshot))
                snapshotPart.FeedData(snapshotStream);
            var snapshotLink = webExtensionPart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                snapshotUrl
            );
            webExtension
                .Element(WE.snapshot)!
                .Add(
                    new XAttribute(R.embed, webExtensionPart.GetIdOfPart(snapshotPart)),
                    new XAttribute(R.link, snapshotLink.Id)
                );
            webExtensionPart.PutXDocument(new XDocument(webExtension));

            var spTree = sourceSlide.GetXDocument().Descendants(P.spTree).Single();
            spTree.Add(
                new XElement(
                    P.graphicFrame,
                    new XElement(
                        P.nvGraphicFramePr,
                        new XElement(P.cNvPr, new XAttribute("id", "1000"), new XAttribute("name", "Add-in")),
                        new XElement(P.cNvGraphicFramePr),
                        new XElement(P.nvPr)
                    ),
                    new XElement(
                        P.xfrm,
                        new XElement(A.off, new XAttribute("x", "0"), new XAttribute("y", "0")),
                        new XElement(A.ext, new XAttribute("cx", "914400"), new XAttribute("cy", "914400"))
                    ),
                    new XElement(
                        A.graphic,
                        new XElement(
                            A.graphicData,
                            new XAttribute("uri", WE.we.NamespaceName),
                            new XElement(
                                WE.webextensionref,
                                new XAttribute(XNamespace.Xmlns + "we", WE.we),
                                new XAttribute(R.id, sourceSlide.GetIdOfPart(webExtensionPart))
                            )
                        )
                    )
                )
            );
            sourceSlide.PutXDocument();
        }

        var sources = new List<SlideSource>
        {
            new(new PmlDocument("PB012-Input.pptx", sourceStream.ToArray()), 0, 1, true),
        };
        var outputPptx = new FileInfo(Path.Combine(TempDir, "PB012-Output.pptx"));
        PresentationBuilder.BuildPresentation(sources).SaveAs(outputPptx.FullName);

        using var pptx = PresentationDocument.Open(outputPptx.FullName, false);
        await Validate(pptx, DocumentFormat.OpenXml.FileFormatVersions.Office2013);

        var slidePart = pptx.PresentationPart!.SlideParts.Single();
        var relId = (string)slidePart.GetXDocument().Descendants(WE.webextensionref).Single().Attribute(R.id);
        var copiedPart = slidePart.WebExtensionParts.Single();
        await Assert.That(slidePart.GetIdOfPart(copiedPart)).IsEqualTo(relId);

        var copiedXml = copiedPart.GetXDocument().Root!;
        await Assert.That((string)copiedXml.Attribute("id")).IsEqualTo((string)webExtension.Attribute("id"));
        await Assert
            .That((string)copiedXml.Element(WE.reference)!.Attribute("id"))
            .IsEqualTo((string)webExtension.Element(WE.reference)!.Attribute("id"));

        var copiedSnapshot = copiedXml.Element(WE.snapshot)!;
        var snapshotId = (string)copiedSnapshot.Attribute(R.embed);
        await Assert.That(ReadAllBytes(copiedPart.GetPartById(snapshotId))).IsEquivalentTo(snapshot);

        var linkedSnapshotId = (string)copiedSnapshot.Attribute(R.link);
        var copiedSnapshotLink = copiedPart.ExternalRelationships.Single(r => r.Id == linkedSnapshotId);
        await Assert.That(copiedSnapshotLink.Uri).IsEqualTo(snapshotUrl);
    }

    private static byte[] ReadAllBytes(OpenXmlPart part)
    {
        using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static string[] GetMediaDataContentTypes(FileInfo fi)
    {
        using var ptDoc = PresentationDocument.Open(fi.FullName, false);
        return ptDoc
            .PresentationPart.SlideParts.SelectMany(p =>
                p.DataPartReferenceRelationships.Select(d => d.DataPart.ContentType)
            )
            .Distinct()
            .OrderBy(m => m)
            .ToArray();
    }
}
