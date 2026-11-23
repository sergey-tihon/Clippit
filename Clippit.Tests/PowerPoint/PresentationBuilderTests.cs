// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit.PowerPoint;
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
