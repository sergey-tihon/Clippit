using System.Xml.Linq;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

public class SlidePartDataTests : TestsBase
{
    private static readonly string SourceDirectory = "../../../../TestFiles/PublishSlides/";
    private static readonly string TestPptxPath = Path.Combine(SourceDirectory, "BRK3066.pptx");

    [Test]
    public async Task SlideLayoutData_SamePartConstructedTwice_ComparesEqual()
    {
        using var doc = PresentationDocument.Open(TestPptxPath, false);
        var layoutPart = doc.PresentationPart!.SlideMasterParts.First().SlideLayoutParts.First();

        var data1 = new SlideLayoutData(layoutPart, 1.0);
        var data2 = new SlideLayoutData(layoutPart, 1.0);

        await Assert.That(data1.CompareTo(data2)).IsEqualTo(0);
    }

    [Test]
    public async Task ThemeData_SamePartConstructedTwice_ComparesEqual()
    {
        using var doc = PresentationDocument.Open(TestPptxPath, false);
        var themePart = doc.PresentationPart!.SlideMasterParts.First().ThemePart!;

        var data1 = new ThemeData(themePart, 1.0);
        var data2 = new ThemeData(themePart, 1.0);

        await Assert.That(data1.CompareTo(data2)).IsEqualTo(0);
    }

    [Test]
    public async Task SlideMasterData_SamePartConstructedTwice_ComparesEqual()
    {
        using var doc = PresentationDocument.Open(TestPptxPath, false);
        var masterPart = doc.PresentationPart!.SlideMasterParts.First();

        var data1 = new SlideMasterData(masterPart, 1.0);
        var data2 = new SlideMasterData(masterPart, 1.0);

        await Assert.That(data1.CompareTo(data2)).IsEqualTo(0);
    }

    /// <summary>
    /// Verifies that constructing SlideLayoutData does not mutate the source XDocument —
    /// specifically that noise attributes (smtClean, dirty, r:id, r:embed) are not stripped
    /// from the original loaded document.
    /// </summary>
    [Test]
    public async Task SlideLayoutData_ConstructionDoesNotMutateSourceXDocument()
    {
        using var doc = PresentationDocument.Open(TestPptxPath, false);
        var layoutPart = doc.PresentationPart!.SlideMasterParts.First().SlideLayoutParts.First();

        // Inject a noise attribute into the layout's spTree before building the descriptor.
        var xDoc = layoutPart.GetXDocument();
        XNamespace pns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        var spTree = xDoc.Descendants(pns + "spTree").First();
        spTree.SetAttributeValue("dirty", "1");

        // Build SlideLayoutData — this triggers GetShapeDescriptor / NormalizeXml.
        _ = new SlideLayoutData(layoutPart, 1.0);

        // The original XDocument must still carry the injected attribute after normalization.
        var afterAttr = xDoc.Descendants(pns + "spTree").First().Attribute("dirty");
        await Assert.That(afterAttr).IsNotNull();
        await Assert.That(afterAttr!.Value).IsEqualTo("1");
    }

    /// <summary>
    /// Two layouts with identical XML but a differing noise attribute (smtClean) must still compare
    /// equal, because the descriptor strips noise attributes before comparing.
    /// </summary>
    [Test]
    public async Task SlideLayoutData_DifferingOnlyInNoiseAttribute_ComparesEqual()
    {
        using var doc = PresentationDocument.Open(TestPptxPath, false);
        var layoutPart = doc.PresentationPart!.SlideMasterParts.First().SlideLayoutParts.First();

        // Build baseline descriptor.
        var data1 = new SlideLayoutData(layoutPart, 1.0);

        // Inject a noise attribute and rebuild. NormalizeXml must strip it before comparing.
        var xDoc = layoutPart.GetXDocument();
        XNamespace pns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        xDoc.Descendants(pns + "spTree").First().SetAttributeValue("smtClean", "0");

        var data2 = new SlideLayoutData(layoutPart, 1.0);

        await Assert.That(data1.CompareTo(data2)).IsEqualTo(0);
    }

    [Test]
    public async Task ScaleShapes_ScalesNumericAttributes()
    {
        XNamespace ans = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var spTree = new XElement(
            "spTree",
            new XElement(ans + "off", new XAttribute("x", "1000"), new XAttribute("y", "2000")),
            new XElement(ans + "ext", new XAttribute("cx", "3000"), new XAttribute("cy", "4000"))
        );

        SlidePartData<SlideLayoutPart>.ScaleShapes(spTree, 2.0);

        await Assert.That(spTree.Descendants(ans + "off").First().Attribute("x")!.Value).IsEqualTo("2000");
        await Assert.That(spTree.Descendants(ans + "off").First().Attribute("y")!.Value).IsEqualTo("4000");
        await Assert.That(spTree.Descendants(ans + "ext").First().Attribute("cx")!.Value).IsEqualTo("6000");
        await Assert.That(spTree.Descendants(ans + "ext").First().Attribute("cy")!.Value).IsEqualTo("8000");
    }

    [Test]
    public async Task ScaleShapes_ScaleFactorOne_LeavesAttributesUnchanged()
    {
        XNamespace ans = "http://schemas.openxmlformats.org/drawingml/2006/main";
        var spTree = new XElement(
            "spTree",
            new XElement(ans + "ext", new XAttribute("cx", "5000"), new XAttribute("cy", "6000"))
        );

        SlidePartData<SlideLayoutPart>.ScaleShapes(spTree, 1.0);

        await Assert.That(spTree.Descendants(ans + "ext").First().Attribute("cx")!.Value).IsEqualTo("5000");
        await Assert.That(spTree.Descendants(ans + "ext").First().Attribute("cy")!.Value).IsEqualTo("6000");
    }
}
