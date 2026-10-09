using System.Xml;
using System.Xml.Linq;
using Clippit.PowerPoint;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

public class SvgToPptxConverterTests : TestsBase
{
    private static readonly XNamespace Presentation = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";

    [Test]
    public async Task SVG001_BasicShapesAndConnectorProduceValidPresentation()
    {
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720">
              <g id="client" data-pptx-kind="node">
                <rect x="100" y="280" width="240" height="100" rx="16" fill="#e8f0fe" stroke="#4285f4" stroke-width="2" />
                <text x="220" y="345" text-anchor="middle" font-size="28" fill="#202124">Client</text>
              </g>
              <g id="api" data-pptx-kind="node">
                <rect x="600" y="280" width="240" height="100" rx="16" fill="#e6f4ea" stroke="#34a853" stroke-width="2" />
                <text x="720" y="345" text-anchor="middle" font-size="28" fill="#202124">API</text>
              </g>
              <line data-pptx-kind="connector" data-pptx-from="client:right" data-pptx-to="api:left" x1="340" y1="330" x2="600" y2="330" stroke="#5f6368" stroke-width="3" />
            </svg>
            """;

        var document = SvgToPptxConverter.Convert(svg);
        var output = Path.Combine(TempDir, "SVG001-basic.pptx");
        document.SaveAs(output);

        using var presentation = PresentationDocument.Open(output, false);
        await Validate(presentation);
        await Assert.That(presentation.PresentationPart!.SlideParts).HasCount().EqualTo(1);
        var slide = presentation.PresentationPart.SlideParts.Single();
        await Assert.That(slide.GetXDocument().Descendants(Presentation + "cxnSp").Count()).IsEqualTo(1);
        await Assert.That(slide.GetXDocument().Descendants(Presentation + "sp").Count()).IsEqualTo(4);
        await Assert.That(slide.GetXDocument().Descendants(Drawing + "stCxn").Count()).IsEqualTo(1);
        await Assert.That(slide.GetXDocument().Descendants(Drawing + "endCxn").Count()).IsEqualTo(1);
    }

    [Test]
    public async Task SVG002_PathAndMarkerBecomeNativeDrawingObjects()
    {
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720">
              <path id="icon" d="M 100 100 L 180 100 L 140 180 Z" fill="#4285f4" stroke="#1a237e" stroke-width="3" />
              <g fill="none" stroke="#202124" marker-end="url(#arrow)">
                <path d="M 220 140 C 280 80 340 200 420 140" />
              </g>
              <defs><marker id="arrow" /></defs>
            </svg>
            """;

        var document = SvgToPptxConverter.Convert(svg);
        var output = Path.Combine(TempDir, "SVG002-path.pptx");
        document.SaveAs(output);

        using var presentation = PresentationDocument.Open(output, false);
        await Validate(presentation);
        var slide = presentation.PresentationPart!.SlideParts.Single();
        var slideXml = slide.GetXDocument();
        await Assert.That(slideXml.Descendants(Drawing + "custGeom")).HasCount().EqualTo(2);
        await Assert.That(slideXml.Descendants(Drawing + "cubicBezTo")).HasCount().EqualTo(1);
        await Assert
            .That(slideXml.Descendants(Drawing + "tailEnd").Any(e => (string?)e.Attribute("type") == "triangle"))
            .IsTrue();
    }

    [Test]
    public async Task SVG003_TemplateContentIsPreservedAndGeneratedShapesAreAppended()
    {
        var template = SvgToPptxConverter.Convert(
            """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720">
              <rect id="template-box" x="20" y="20" width="180" height="80" fill="#eeeeee" />
              <text x="110" y="70" text-anchor="middle">Template</text>
            </svg>
            """
        );

        var document = SvgToPptxConverter.Convert(
            """
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1280 720">
              <rect id="generated-box" x="300" y="20" width="180" height="80" fill="#e8f0fe" />
              <text x="390" y="70" text-anchor="middle">Generated</text>
            </svg>
            """,
            new SvgToPptxConverterSettings { Template = template }
        );

        using var stream = new MemoryStream(document.DocumentByteArray);
        using var presentation = PresentationDocument.Open(stream, false);
        await Validate(presentation);
        var slide = presentation.PresentationPart!.SlideParts.Single();
        var slideXml = slide.GetXDocument();
        var text = string.Concat(slideXml.Descendants(Drawing + "t").Select(element => element.Value));

        await Assert.That(text).Contains("Template");
        await Assert.That(text).Contains("Generated");
        await Assert.That(slideXml.Descendants(Presentation + "sp").Count()).IsEqualTo(4);
    }

    [Test]
    public async Task SVG004_MultipleSourcesProduceMultipleSlides()
    {
        var document = SvgToPptxConverter.Convert([
            new SvgSlideSource(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1280 720\"><text x=\"10\" y=\"30\">One</text></svg>"
            ),
            new SvgSlideSource(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1280 720\"><text x=\"10\" y=\"30\">Two</text></svg>"
            ),
        ]);

        using var stream = new MemoryStream(document.DocumentByteArray);
        using var presentation = PresentationDocument.Open(stream, false);
        await Assert.That(presentation.PresentationPart!.SlideParts).HasCount().EqualTo(2);
        var text = presentation.PresentationPart.SlideParts.SelectMany(slide =>
            slide.GetXDocument().Descendants(Drawing + "t")
        );
        await Assert.That(text.Select(element => element.Value)).Contains("One");
        await Assert.That(text.Select(element => element.Value)).Contains("Two");
    }

    [Test]
    public async Task SVG005_TemplateSlideSizeIsUsedAndExistingContentIsPreserved()
    {
        var template = SvgToPptxConverter.Convert(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1280 720\"><text x=\"10\" y=\"30\">Template</text></svg>"
        );
        using var templateStream = new MemoryStream(template.DocumentByteArray);
        using (var templatePresentation = PresentationDocument.Open(templateStream, true))
        {
            var presentationXml = templatePresentation.PresentationPart!.GetXDocument();
            presentationXml.Root!.Element(Presentation + "sldSz")!.SetAttributeValue("cx", 10_000_000);
            presentationXml.Root!.Element(Presentation + "sldSz")!.SetAttributeValue("cy", 5_000_000);
            templatePresentation.PresentationPart.PutXDocument(presentationXml);
        }

        var sizedTemplate = new PmlDocument("template.pptx", templateStream.ToArray());
        var document = SvgToPptxConverter.Convert(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1280 720\"><rect x=\"0\" y=\"0\" width=\"1280\" height=\"720\" fill=\"#eeeeee\" /></svg>",
            new SvgToPptxConverterSettings { Template = sizedTemplate }
        );

        using var outputStream = new MemoryStream(document.DocumentByteArray);
        using var output = PresentationDocument.Open(outputStream, false);
        await Validate(output);
        var slideXml = output.PresentationPart!.SlideParts.Single().GetXDocument();
        var generatedTransform = slideXml.Descendants(Drawing + "xfrm").Last();
        // The template is wider than the diagram, so the diagram is fitted, centered and never stretched.
        var scale = Math.Min(10_000_000d / 1280d, 5_000_000d / 720d);
        var expectedWidth = (long)Math.Round(1280 * scale);
        await Assert.That((long)generatedTransform.Element(Drawing + "ext")!.Attribute("cx")!).IsEqualTo(expectedWidth);
        await Assert.That((long)generatedTransform.Element(Drawing + "ext")!.Attribute("cy")!).IsEqualTo(5_000_000);
        var offsetX = (long)generatedTransform.Element(Drawing + "off")!.Attribute("x")!;
        await Assert.That(Math.Abs(offsetX - (10_000_000 - expectedWidth) / 2)).IsLessThanOrEqualTo(1);
        await Assert.That(slideXml.Descendants(Drawing + "t").Select(element => element.Value)).Contains("Template");
    }

    [Test]
    public async Task SVG006_ReversedLinesRemainValid()
    {
        var document = SvgToPptxConverter.Convert(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1280 720\"><line x1=\"800\" y1=\"400\" x2=\"200\" y2=\"100\" stroke=\"#000000\" /></svg>"
        );

        using var stream = new MemoryStream(document.DocumentByteArray);
        using var presentation = PresentationDocument.Open(stream, false);
        await Validate(presentation);
        var xfrm = presentation
            .PresentationPart!.SlideParts.Single()
            .GetXDocument()
            .Descendants(Drawing + "xfrm")
            .Last();
        await Assert.That((long)xfrm.Element(Drawing + "ext")!.Attribute("cx")!).IsGreaterThan(0);
        await Assert.That((long)xfrm.Element(Drawing + "ext")!.Attribute("cy")!).IsGreaterThan(0);
    }

    [Test]
    public async Task SVG007_RejectsUnsafeAndInvalidInput()
    {
        const string unsafeSvg =
            "<!DOCTYPE svg [<!ENTITY xxe SYSTEM 'file:///etc/passwd'>]><svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1280 720\"><text>&xxe;</text></svg>";
        await Assert.That(() => SvgToPptxConverter.Convert(unsafeSvg)).Throws<XmlException>();

        const string invalidViewBox = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 NaN 720\" />";
        await Assert.That(() => SvgToPptxConverter.Convert(invalidViewBox)).Throws<FormatException>();
    }

    [Test]
    public async Task SVG008_StandardSvgDoctypeIsAcceptedWithoutProcessingEntities()
    {
        const string svg =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><!DOCTYPE svg PUBLIC \"-//W3C//DTD SVG 1.1//EN\" \"http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd\"><svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1280 720\"><text x=\"10\" y=\"30\">SVG</text></svg>";

        var document = SvgToPptxConverter.Convert(svg);
        await Assert.That(document.DocumentByteArray).IsNotEmpty();
    }

    [Test]
    public async Task SVG009_DrawioHtmlLabelsBecomeTextAndTheNotSvgBannerIsIgnored()
    {
        const string svg = """
            <svg xmlns="http://www.w3.org/2000/svg" xmlns:xhtml="http://www.w3.org/1999/xhtml" viewBox="0 0 982 1204">
              <g transform="translate(0.5,0.5)">
                <rect x="329.67" y="28" width="50" height="50" fill="#e8f1fb" stroke="#1f6feb" />
              </g>
              <g>
                <g>
                  <switch>
                    <foreignObject style="overflow: visible; text-align: left;" pointer-events="none" width="100%" height="100%">
                      <xhtml:div style="display: flex; padding-top: 58px; margin-left: 354.67px;">
                        <xhtml:div style="box-sizing: border-box; font-size: 0; text-align: center; color: #232F3E;">
                          <xhtml:div style="display: inline-block; font-size: 10px; font-family: Helvetica; color: light-dark(#232F3E, #bdc7d4);">Market user</xhtml:div>
                        </xhtml:div>
                      </xhtml:div>
                    </foreignObject>
                  </switch>
                </g>
              </g>
              <switch>
                <g requiredFeatures="http://www.w3.org/TR/SVG11/feature#Extensibility" />
                <a transform="translate(0,-5)" xlink:href="https://www.drawio.com/doc/faq/svg-export-text-problems" xmlns:xlink="http://www.w3.org/1999/xlink">
                  <text text-anchor="middle" font-size="10px" x="50%" y="100%">Text is not SVG - cannot display</text>
                </a>
              </switch>
            </svg>
            """;

        var document = SvgToPptxConverter.Convert(svg);
        var output = Path.Combine(TempDir, "SVG009-drawio.pptx");
        document.SaveAs(output);

        using var presentation = PresentationDocument.Open(output, false);
        await Validate(presentation);
        var texts = presentation
            .PresentationPart!.SlideParts.Single()
            .GetXDocument()
            .Descendants(Drawing + "t")
            .Select(t => (string?)t)
            .ToArray();
        await Assert.That(texts).IsEquivalentTo(["Market user"]);
    }

    [Test]
    public async Task SVG010_DiagramIsFittedIntoTheDefaultSlideWithoutChangingItsSize()
    {
        const string svg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 1000 2000\"><rect x=\"0\" y=\"0\" width=\"1000\" height=\"2000\" fill=\"#eeeeee\" /></svg>";

        var document = SvgToPptxConverter.Convert(svg);
        var output = Path.Combine(TempDir, "SVG010-fit.pptx");
        document.SaveAs(output);

        using var presentation = PresentationDocument.Open(output, false);
        await Validate(presentation);
        var presentationXml = presentation.PresentationPart!.GetXDocument();
        var slideSize = presentationXml.Root!.Element(Presentation + "sldSz")!;
        await Assert.That((long)slideSize.Attribute("cx")!).IsEqualTo(12_192_000);
        await Assert.That((long)slideSize.Attribute("cy")!).IsEqualTo(6_858_000);

        // 1000x2000 fits by height: 3429 EMU per unit, centered horizontally, never stretched.
        var transform = presentation
            .PresentationPart.SlideParts.Single()
            .GetXDocument()
            .Descendants(Drawing + "xfrm")
            .Last();
        await Assert.That((long)transform.Element(Drawing + "ext")!.Attribute("cx")!).IsEqualTo(3_429_000);
        await Assert.That((long)transform.Element(Drawing + "ext")!.Attribute("cy")!).IsEqualTo(6_858_000);
        await Assert.That((long)transform.Element(Drawing + "off")!.Attribute("x")!).IsEqualTo(4_381_500);
        await Assert.That((long)transform.Element(Drawing + "off")!.Attribute("y")!).IsEqualTo(0);
    }
}
