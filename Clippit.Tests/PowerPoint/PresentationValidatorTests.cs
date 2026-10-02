using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Clippit;
using Clippit.Core;
using Clippit.PowerPoint;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.PowerPoint;

#pragma warning disable CA1707 // Test names follow the repository's prefix_code_descriptive_name convention.
public sealed class PresentationValidatorTests
{
    private static readonly DirectoryInfo s_testFiles = new("../../../../TestFiles");

    [Test]
    public async Task PV001_ValidDeck_HasNoDiagnostics()
    {
        await using var stream = File.OpenRead(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsTrue();
        await Assert.That(result.Diagnostics).IsEmpty();
    }

    [Test]
    public async Task PV002_NonPptxStream_ReturnsPackageDiagnostic()
    {
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes("not a zip package"));

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert.That(result.Diagnostics).HasCount(1);
        await Assert.That(result.Diagnostics[0].Kind).IsEqualTo(OpenXmlValidationDiagnosticKinds.Package);
    }

    [Test]
    public async Task PV003_DanglingRelationship_ReturnsRelationshipDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        InjectDanglingOleObjRelId(stream, "rId_library_validator_dangling");
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert
            .That(
                result.Diagnostics.Any(d =>
                    d.Kind == OpenXmlValidationDiagnosticKinds.Relationship
                    && d.RelationshipId == "rId_library_validator_dangling"
                )
            )
            .IsTrue();
    }

    [Test]
    public async Task PV004_MalformedSectionList_ReturnsPptxSectionDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        AddMalformedSectionList(stream);
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert
            .That(result.Diagnostics.Any(d => d.Kind == PresentationValidationDiagnosticKinds.Section))
            .IsTrue();
    }

    [Test]
    public async Task PV005_RelationshipValidator_DoesNotAnnotateParts()
    {
        await using var stream = File.OpenRead(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        using var document = PresentationDocument.Open(stream, false);
        var slidePart = document.PresentationPart!.SlideParts.First();

        _ = RelationshipValidator.Validate(document);

        await Assert.That(slidePart.Annotation<XDocument>()).IsNull();
    }

    [Test]
    public async Task PV006_MalformedXmlPart_ReturnsPackageDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        ReplaceSlideXml(stream, "<broken");
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert.That(result.Diagnostics.Any(d => d.Kind == OpenXmlValidationDiagnosticKinds.Package)).IsTrue();
    }

    [Test]
    public async Task PV007_BracelessSectionGuid_ReturnsPptxSectionDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        AddMalformedSectionList(stream, useBracelessSectionId: true);
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert
            .That(
                result.Diagnostics.Any(d =>
                    d.Kind == PresentationValidationDiagnosticKinds.Section && d.Attribute == "id"
                )
            )
            .IsTrue();
    }

    [Test]
    public async Task PV008_MalformedPresentationXml_ReturnsPackageDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        ReplaceZipEntryText(stream, "ppt/presentation.xml", "<broken");
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert.That(result.Diagnostics.Any(d => d.Kind == OpenXmlValidationDiagnosticKinds.Package)).IsTrue();
    }

    [Test]
    public async Task PV009_MalformedXmlPart_OpenDocumentValidation_ReturnsPackageDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        ReplaceSlideXml(stream, "<broken");
        stream.Position = 0;
        using var document = PresentationDocument.Open(stream, false);

        var result = PresentationValidator.Validate(document);

        await Assert.That(result.Valid).IsFalse();
        await Assert.That(result.Diagnostics.Any(d => d.Kind == OpenXmlValidationDiagnosticKinds.Package)).IsTrue();
    }

    [Test]
    public async Task PV010_SectionSlideReferencedByRelationshipId_ReturnsPptxSectionDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        AddSectionWithSlideReference(stream, useRelationshipId: true);
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert
            .That(
                result.Diagnostics.Any(d =>
                    d.Kind == PresentationValidationDiagnosticKinds.Section
                    && d.Element == "sldId"
                    && d.Description.Contains("references slide by relationship id", StringComparison.Ordinal)
                )
            )
            .IsTrue();
    }

    [Test]
    public async Task PV011_SectionSlideMissingNumericId_ReturnsPptxSectionDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        AddSectionWithSlideReference(stream, omitNumericId: true);
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert
            .That(
                result.Diagnostics.Any(d =>
                    d.Kind == PresentationValidationDiagnosticKinds.Section
                    && d.Attribute == "id"
                    && d.Description.Contains("without numeric id", StringComparison.Ordinal)
                )
            )
            .IsTrue();
    }

    [Test]
    public async Task PV012_SectionSlideNonNumericId_ReturnsPptxSectionDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        AddSectionWithSlideReference(stream, numericId: "not-a-number");
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert
            .That(
                result.Diagnostics.Any(d =>
                    d.Kind == PresentationValidationDiagnosticKinds.Section
                    && d.Description.Contains("non-numeric slide id", StringComparison.Ordinal)
                )
            )
            .IsTrue();
    }

    [Test]
    public async Task PV013_SectionSlideIdNotInSldIdLst_ReturnsPptxSectionDiagnostic()
    {
        using var stream = CreateExpandableStream(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"));
        AddSectionWithSlideReference(stream, numericId: "999999");
        stream.Position = 0;

        var result = PresentationValidator.Validate(stream);

        await Assert.That(result.Valid).IsFalse();
        await Assert
            .That(
                result.Diagnostics.Any(d =>
                    d.Kind == PresentationValidationDiagnosticKinds.Section
                    && d.Description.Contains("not present in p:sldIdLst", StringComparison.Ordinal)
                )
            )
            .IsTrue();
    }

    private static MemoryStream CreateExpandableStream(string path)
    {
        var stream = new MemoryStream();
        using (var file = File.OpenRead(path))
            file.CopyTo(stream);
        stream.Position = 0;
        return stream;
    }

    private static void InjectDanglingOleObjRelId(Stream pptxStream, string danglingId)
    {
        XNamespace pNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace rNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var zip = new ZipArchive(pptxStream, ZipArchiveMode.Update, leaveOpen: true);
        var slideEntry = zip.Entries.First(e =>
            e.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
            && e.FullName.EndsWith(".xml", StringComparison.Ordinal)
        );

        XDocument xDoc;
        using (var stream = slideEntry.Open())
            xDoc = XDocument.Load(stream);

        var target = xDoc.Descendants(pNs + "spTree").FirstOrDefault() ?? xDoc.Root;
        target?.Add(new XElement(pNs + "oleObj", new XAttribute(rNs + "id", danglingId)));

        var fullName = slideEntry.FullName;
        slideEntry.Delete();
        var newEntry = zip.CreateEntry(fullName);
        using var writer = new StreamWriter(newEntry.Open());
        using var xmlWriter = System.Xml.XmlWriter.Create(writer);
        xDoc.WriteTo(xmlWriter);
    }

    private static void ReplaceSlideXml(Stream pptxStream, string xml)
    {
        using var zip = new ZipArchive(pptxStream, ZipArchiveMode.Update, leaveOpen: true);
        var slideEntry = zip.Entries.First(e =>
            e.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal)
            && e.FullName.EndsWith(".xml", StringComparison.Ordinal)
        );

        var fullName = slideEntry.FullName;
        slideEntry.Delete();
        var newEntry = zip.CreateEntry(fullName);
        using var writer = new StreamWriter(newEntry.Open());
        writer.Write(xml);
    }

    private static void ReplaceZipEntryText(Stream pptxStream, string entryName, string text)
    {
        using var zip = new ZipArchive(pptxStream, ZipArchiveMode.Update, leaveOpen: true);
        var entry = zip.GetEntry(entryName)!;
        entry.Delete();
        var newEntry = zip.CreateEntry(entryName);
        using var writer = new StreamWriter(newEntry.Open());
        writer.Write(text);
    }

    private static void AddMalformedSectionList(Stream pptxStream, bool useBracelessSectionId = false)
    {
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace p14 = "http://schemas.microsoft.com/office/powerpoint/2010/main";
        XNamespace r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        using var zip = new ZipArchive(pptxStream, ZipArchiveMode.Update, leaveOpen: true);
        var presentationEntry = zip.GetEntry("ppt/presentation.xml")!;

        XDocument xDoc;
        using (var stream = presentationEntry.Open())
            xDoc = XDocument.Load(stream);

        var firstSlide = xDoc.Root!.Element(p + "sldIdLst")!.Elements(p + "sldId").First();
        var relId = (string)firstSlide.Attribute(r + "id")!;
        xDoc.Root.Add(
            new XElement(
                p14 + "sectionLst",
                new XElement(
                    p14 + "section",
                    new XAttribute(NoNamespace.name, "Malformed"),
                    new XAttribute(
                        NoNamespace.id,
                        useBracelessSectionId ? Guid.NewGuid().ToString("D") : Guid.NewGuid().ToString("B")
                    ),
                    new XElement(p14 + "sldIdLst", new XElement(p14 + "sldId", new XAttribute(r + "id", relId)))
                )
            )
        );

        presentationEntry.Delete();
        var newEntry = zip.CreateEntry("ppt/presentation.xml");
        using var writer = new StreamWriter(newEntry.Open());
        using var xmlWriter = System.Xml.XmlWriter.Create(writer);
        xDoc.WriteTo(xmlWriter);
    }

    /// <summary>
    /// Adds a well-formed section with a single slide reference whose id attributes can be
    /// individually tweaked to exercise each PresentationSectionValidator error branch.
    /// </summary>
    private static void AddSectionWithSlideReference(
        Stream pptxStream,
        bool useRelationshipId = false,
        bool omitNumericId = false,
        string? numericId = null
    )
    {
        var p = XNamespace.Get("http://schemas.openxmlformats.org/presentationml/2006/main");
        var p14 = XNamespace.Get("http://schemas.microsoft.com/office/powerpoint/2010/main");
        var r = XNamespace.Get("http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        using var zip = new ZipArchive(pptxStream, ZipArchiveMode.Update, leaveOpen: true);
        var presentationEntry = zip.GetEntry("ppt/presentation.xml")!;

        XDocument xDoc;
        using (var stream = presentationEntry.Open())
            xDoc = XDocument.Load(stream);

        var firstSlide = xDoc.Root!.Element(p + "sldIdLst")!.Elements(p + "sldId").First();
        var relId = (string)firstSlide.Attribute(r + "id")!;
        var firstSlideId = (string)firstSlide.Attribute(NoNamespace.id)!;

        var sldIdElement = new XElement(p14 + "sldId");
        if (useRelationshipId)
            sldIdElement.SetAttributeValue(r + "id", relId);
        if (!omitNumericId)
            sldIdElement.SetAttributeValue(NoNamespace.id, numericId ?? firstSlideId);

        xDoc.Root.Add(
            new XElement(
                p14 + "sectionLst",
                new XElement(
                    p14 + "section",
                    new XAttribute(NoNamespace.name, "Section1"),
                    new XAttribute(NoNamespace.id, Guid.NewGuid().ToString("B")),
                    new XElement(p14 + "sldIdLst", sldIdElement)
                )
            )
        );

        presentationEntry.Delete();
        var newEntry = zip.CreateEntry("ppt/presentation.xml");
        using var writer = new StreamWriter(newEntry.Open());
        using var xmlWriter = System.Xml.XmlWriter.Create(writer);
        xDoc.WriteTo(xmlWriter);
    }
}
#pragma warning restore CA1707
