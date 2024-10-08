using System;
using System.IO;
using System.Threading.Tasks;
using System.Xml.Linq;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;
using Xunit;

namespace Clippit.Tests.PowerPoint;

public partial class PresentationBuilderSlidePublishingTests
{
    [Theory]
    [ClassData(typeof(PublishingTestData))]
    public async Task PublishUsingMemDocs(string sourcePath)
    {
        var fileName = Path.GetFileNameWithoutExtension(sourcePath);
        var targetDir = Path.Combine(TargetDirectory, fileName);
        if (Directory.Exists(targetDir))
            Directory.Delete(targetDir, true);
        Directory.CreateDirectory(targetDir);

        await using var srcStream = File.Open(sourcePath, FileMode.Open);
        var openSettings = new OpenSettings { AutoSave = false };
        using var srcDoc = OpenXmlExtensions.OpenPresentation(srcStream, false, openSettings);

        var slideNumber = 0;
        var slidesIds = PresentationBuilderTools.GetSlideIdsInOrder(srcDoc);
        foreach (var slideId in slidesIds)
        {
            var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slideId);
            var title = PresentationBuilderTools.GetSlideTitle(srcSlidePart.GetXElement());

            using var ms = new MemoryStream();
            using var streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument(ms);
            using (var output = streamDoc.GetPresentationDocument(new OpenSettings { AutoSave = false }))
            {
                using (var builder = PresentationBuilder.Create(output))
                {
                    try
                    {
                        var newSlidePart = builder.AddSlidePart(srcSlidePart);

                        // Remove the show attribute from the slide element (if it exists)
                        var slideDocument = newSlidePart.GetXDocument();
                        slideDocument.Root?.Attribute(NoNamespace.show)?.Remove();
                    }
                    catch (PresentationBuilderInternalException dbie)
                    {
                        if (dbie.Message.Contains("{0}"))
                            throw new PresentationBuilderException(string.Format(dbie.Message, srcSlidePart.Uri));
                        throw;
                    }
                }

                // Set the title of the new presentation to the title of the slide
                output.PackageProperties.Title = title;
            }

            streamDoc.ClosePackage();

            var slideFileName = string.Concat(fileName, $"_{++slideNumber:000}.pptx");
            await using var fs = File.Create(Path.Combine(TargetDirectory, slideFileName));
            ms.Position = 0;
            await ms.CopyToAsync(fs);

            srcSlidePart.RemoveAnnotations<XDocument>();
            srcSlidePart.UnloadRootElement();
        }

        Log.WriteLine($"GC Total Memory: {GC.GetTotalMemory(false) / 1024 / 1024} MB");
    }
}
