using System.Text.RegularExpressions;
using System.Xml.Linq;
using Clippit.PowerPoint;

namespace Clippit.Tests.PowerPoint.Samples
{
    public class PresentationBuilderSamples() : Clippit.Tests.TestsBase
    {
        private static string GetFilePath(string path) =>
            Path.Combine("../../../PowerPoint/Samples/PresentationBuilder/", path);

        [Test]
        public void Sample1()
        {
            var source1 = GetFilePath("Sample1/Contoso.pptx");
            var source2 = GetFilePath("Sample1/Companies.pptx");
            var source3 = GetFilePath("Sample1/Customer Content.pptx");
            var source4 = GetFilePath("Sample1/Presentation One.pptx");
            var source5 = GetFilePath("Sample1/Presentation Two.pptx");
            var source6 = GetFilePath("Sample1/Presentation Three.pptx");
            var contoso1 = GetFilePath("Sample1/Contoso One.pptx");
            var contoso2 = GetFilePath("Sample1/Contoso Two.pptx");
            var contoso3 = GetFilePath("Sample1/Contoso Three.pptx");
            var sourceDoc = new PmlDocument(source1);
            var sources = new List<SlideSource>
            {
                new(sourceDoc, 0, 1, false), // Title
                new(sourceDoc, 1, 1, false), // First intro (of 3)
                new(sourceDoc, 4, 2, false), // Sales bios
                new(sourceDoc, 9, 3, false), // Content slides
                new(sourceDoc, 13, 1, false), // Closing summary
            };
            PresentationBuilder.BuildPresentation(sources).SaveAs(Path.Combine(TempDir, "Out1.pptx"));
            sources =
            [
                new SlideSource(new PmlDocument(source2), 2, 1, true), // Choose company
                new SlideSource(new PmlDocument(source3), false), // Content
            ];
            PresentationBuilder.BuildPresentation(sources).SaveAs(Path.Combine(TempDir, "Out2.pptx"));
            sources =
            [
                new SlideSource(new PmlDocument(source4), true),
                new SlideSource(new PmlDocument(source5), true),
                new SlideSource(new PmlDocument(source6), true),
            ];
            PresentationBuilder.BuildPresentation(sources).SaveAs(Path.Combine(TempDir, "Out3.pptx"));
            sources =
            [
                new SlideSource(new PmlDocument(contoso1), true),
                new SlideSource(new PmlDocument(contoso2), true),
                new SlideSource(new PmlDocument(contoso3), true),
            ];
            PresentationBuilder.BuildPresentation(sources).SaveAs(Path.Combine(TempDir, "Out4.pptx"));
        }

        [Test]
        public void Sample2()
        {
            var presentation = GetFilePath("Sample2/Presentation1.pptx");
            var hiddenPresentation = GetFilePath("Sample2/HiddenPresentation.pptx");
            // First, load both presentations into byte arrays, simulating retrieving presentations from some source
            // such as a SharePoint server
            var baPresentation = File.ReadAllBytes(presentation);
            var baHiddenPresentation = File.ReadAllBytes(hiddenPresentation);
            // Next, replace "thee" with "the" in the main presentation
            var pmlMainPresentation = new PmlDocument("Main.pptx", baPresentation);
            PmlDocument modifiedMainPresentation;
            using (var streamDoc = new OpenXmlMemoryStreamDocument(pmlMainPresentation))
            {
                using (var document = streamDoc.GetPresentationDocument())
                {
                    var pXDoc = document.PresentationPart.GetXDocument();
                    foreach (var slideId in pXDoc.Root.Elements(P.sldIdLst).Elements(P.sldId))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        var slidePart = document.PresentationPart.GetPartById(slideRelId);
                        var slideXDoc = slidePart.GetXDocument();
                        var paragraphs = slideXDoc.Descendants(A.p).ToList();
                        OpenXmlRegex.Replace(paragraphs, new Regex("thee"), "the", null);
                        slidePart.PutXDocument();
                    }
                }

                modifiedMainPresentation = streamDoc.GetModifiedPmlDocument();
            }

            // Combine the two presentations into a single presentation
            var slideSources = new List<SlideSource>
            {
                new(modifiedMainPresentation, 0, 1, true),
                new(new PmlDocument("Hidden.pptx", baHiddenPresentation), true),
                new(modifiedMainPresentation, 1, true),
            };
            var combinedPresentation = PresentationBuilder.BuildPresentation(slideSources);
            // Replace <# TRADEMARK #> with AdventureWorks (c)
            PmlDocument modifiedCombinedPresentation;
            using (var streamDoc = new OpenXmlMemoryStreamDocument(combinedPresentation))
            {
                using (var document = streamDoc.GetPresentationDocument())
                {
                    var pXDoc = document.PresentationPart.GetXDocument();
                    foreach (var slideId in pXDoc.Root.Elements(P.sldIdLst).Elements(P.sldId).Skip(1).Take(1))
                    {
                        var slideRelId = (string)slideId.Attribute(R.id);
                        var slidePart = document.PresentationPart.GetPartById(slideRelId);
                        var slideXDoc = slidePart.GetXDocument();
                        var paragraphs = slideXDoc.Descendants(A.p).ToList();
                        OpenXmlRegex.Replace(paragraphs, new Regex("<# TRADEMARK #>"), "AdventureWorks (c)", null);
                        slidePart.PutXDocument();
                    }
                }

                modifiedCombinedPresentation = streamDoc.GetModifiedPmlDocument();
            }

            // we now have a PmlDocument (which is essentially a byte array) that can be saved as necessary.
            modifiedCombinedPresentation.SaveAs(Path.Combine(TempDir, "ModifiedCombinedPresentation.pptx"));
        }
    }
}
