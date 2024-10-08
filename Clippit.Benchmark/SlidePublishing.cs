using System.Reflection;
using System.Xml.Linq;
using BenchmarkDotNet.Attributes;
using Clippit.PowerPoint;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Benchmark;

[MemoryDiagnoser]
public class SlidePublishing
{
    private const string SourcePath = "TestFiles/PublishSlides/BRK3066.pptx";
    private const string TargetDirectory = "TestFiles/PublishSlides/output";

    [GlobalSetup]
    public void Setup()
    {
        var root = Assembly.GetExecutingAssembly().Location;
        while (!Directory.Exists(Path.Combine(root, ".git")))
            root = Directory.GetParent(root).FullName;
        _filePath = Path.Combine(root, SourcePath);

        if (!File.Exists(_filePath))
            throw new FileNotFoundException("Source file not found", _filePath);

        _pmlDir = Path.Combine(root, TargetDirectory, nameof(PmlDocument));
        _fluentDir = Path.Combine(root, TargetDirectory, nameof(FluentBuilder));

        foreach (var dir in new[] { _pmlDir, _fluentDir })
        {
            if (Directory.Exists(dir))
                Directory.Delete(dir, true);
            Directory.CreateDirectory(dir);
        }
    }

    private string _filePath;
    private string _pmlDir;
    private string _fluentDir;

    [Benchmark]
    public void PmlDocument()
    {
        var fileName = Path.GetFileNameWithoutExtension(_filePath);
        using var srcStream = File.Open(_filePath, FileMode.Open);
        var openSettings = new OpenSettings { AutoSave = false };
        using var srcDoc = PresentationDocument.Open(srcStream, false, openSettings);

        foreach (var slide in PresentationBuilder.PublishSlides(srcDoc, fileName))
        {
            slide.SaveAs(Path.Combine(_pmlDir, Path.GetFileName(slide.FileName)));
        }
    }

    [Benchmark]
    public async Task FluentBuilder()
    {
        var fileName = Path.GetFileNameWithoutExtension(_filePath);

        await using var srcStream = File.Open(_filePath, FileMode.Open);
        var openSettings = new OpenSettings { AutoSave = false };
        using var srcDoc = PresentationDocument.Open(srcStream, false, openSettings);

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
            await using var fs = File.Create(Path.Combine(_fluentDir, slideFileName));
            ms.Position = 0;
            await ms.CopyToAsync(fs);

            srcSlidePart.RemoveAnnotations<XDocument>();
            srcSlidePart.UnloadRootElement();
        }
    }
}
