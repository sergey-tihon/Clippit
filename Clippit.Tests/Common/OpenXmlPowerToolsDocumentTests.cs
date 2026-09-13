// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Excel;
using Clippit.PowerPoint;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="OpenXmlPowerToolsDocument"/>, <see cref="OpenXmlMemoryStreamDocument"/>,
/// <see cref="WmlDocument"/>, <see cref="SmlDocument"/>, and <see cref="PmlDocument"/>.
/// </summary>
public class OpenXmlPowerToolsDocumentTests : TestsBase
{
    private static readonly string TestFilesDir = "../../../../TestFiles/";
    private static readonly string DocxPath = Path.Combine(TestFilesDir, "Blank-wml.docx");
    private static readonly string XlsxPath = Path.Combine(TestFilesDir, "SH001-Table.xlsx");
    private static readonly string PptxPath = Path.Combine(TestFilesDir, "PB001-Input1.pptx");

    // ── OpenXmlPowerToolsDocument.GetDocumentType ───────────────────────────

    [Test]
    public async Task OXD001_GetDocumentType_Docx_ReturnsWordprocessingDocument()
    {
        var doc = new WmlDocument(DocxPath);
        await Assert.That(doc.GetDocumentType()).IsEqualTo(typeof(WordprocessingDocument));
    }

    [Test]
    public async Task OXD002_GetDocumentType_Xlsx_ReturnsSpreadsheetDocument()
    {
        var doc = new SmlDocument(XlsxPath);
        await Assert.That(doc.GetDocumentType()).IsEqualTo(typeof(SpreadsheetDocument));
    }

    [Test]
    public async Task OXD003_GetDocumentType_Pptx_ReturnsPresentationDocument()
    {
        var doc = new PmlDocument(PptxPath);
        await Assert.That(doc.GetDocumentType()).IsEqualTo(typeof(PresentationDocument));
    }

    // ── OpenXmlPowerToolsDocument.FromFileName ──────────────────────────────

    [Test]
    public async Task OXD004_FromFileName_Docx_ReturnsWmlDocument()
    {
        var doc = OpenXmlPowerToolsDocument.FromFileName(DocxPath);
        await Assert.That(doc).IsTypeOf<WmlDocument>();
    }

    [Test]
    public async Task OXD005_FromFileName_Xlsx_ReturnsSmlDocument()
    {
        var doc = OpenXmlPowerToolsDocument.FromFileName(XlsxPath);
        await Assert.That(doc).IsTypeOf<SmlDocument>();
    }

    [Test]
    public async Task OXD006_FromFileName_Pptx_ReturnsPmlDocument()
    {
        var doc = OpenXmlPowerToolsDocument.FromFileName(PptxPath);
        await Assert.That(doc).IsTypeOf<PmlDocument>();
    }

    // ── OpenXmlPowerToolsDocument.FromDocument ──────────────────────────────

    [Test]
    public async Task OXD007_FromDocument_WmlDocument_ReturnsWmlDocument()
    {
        var original = new WmlDocument(DocxPath);
        var copy = OpenXmlPowerToolsDocument.FromDocument(original);
        await Assert.That(copy).IsTypeOf<WmlDocument>();
    }

    [Test]
    public async Task OXD008_FromDocument_SmlDocument_ReturnsSmlDocument()
    {
        var original = new SmlDocument(XlsxPath);
        var copy = OpenXmlPowerToolsDocument.FromDocument(original);
        await Assert.That(copy).IsTypeOf<SmlDocument>();
    }

    [Test]
    public async Task OXD009_FromDocument_PmlDocument_ReturnsPmlDocument()
    {
        var original = new PmlDocument(PptxPath);
        var copy = OpenXmlPowerToolsDocument.FromDocument(original);
        await Assert.That(copy).IsTypeOf<PmlDocument>();
    }

    // ── Copy constructor ────────────────────────────────────────────────────

    [Test]
    public async Task OXD010_CopyConstructor_PreservesFileName()
    {
        var original = new WmlDocument(DocxPath);
        var copy = new WmlDocument(original);
        await Assert.That(copy.FileName).IsEqualTo(original.FileName);
        await Assert.That(copy.DocumentByteArray).IsNotNull();
        await Assert.That(copy.DocumentByteArray.Length).IsGreaterThan(0);
    }

    [Test]
    public async Task OXD011_CopyConstructor_ByteArrayIsIndependentCopy()
    {
        var original = new WmlDocument(DocxPath);
        var copy = new WmlDocument(original);
        // Mutating the copy must not affect the original
        copy.DocumentByteArray[0] ^= 0xFF;
        await Assert.That(copy.DocumentByteArray[0]).IsNotEqualTo(original.DocumentByteArray[0]);
    }

    // ── WriteByteArray ──────────────────────────────────────────────────────

    [Test]
    public async Task OXD012_WriteByteArray_WritesAllBytes()
    {
        var doc = new WmlDocument(DocxPath);
        using var ms = new MemoryStream();
        doc.WriteByteArray(ms);
        await Assert.That(ms.ToArray().Length).IsEqualTo(doc.DocumentByteArray.Length);
        await Assert.That(ms.ToArray().SequenceEqual(doc.DocumentByteArray)).IsTrue();
    }

    // ── SaveAs ──────────────────────────────────────────────────────────────

    [Test]
    public async Task OXD013_SaveAs_CreatesFileWithCorrectContent()
    {
        var doc = new WmlDocument(DocxPath);
        var destPath = Path.Combine(TempDir, "OXD013-saved.docx");
        doc.SaveAs(destPath);
        var savedBytes = File.ReadAllBytes(destPath);
        await Assert.That(savedBytes).IsEquivalentTo(doc.DocumentByteArray);
    }

    // ── GetName ─────────────────────────────────────────────────────────────

    [Test]
    public async Task OXD014_GetName_ReturnsFileNameWithoutPath()
    {
        var doc = new WmlDocument(DocxPath);
        var name = doc.GetName();
        await Assert.That(name).IsEqualTo(Path.GetFileName(DocxPath));
    }

    // ── OpenXmlMemoryStreamDocument factory methods ─────────────────────────

    [Test]
    public async Task OXD020_CreateWordprocessingDocument_ReturnsWordDoc()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
        await Assert.That(memDoc.GetDocumentType()).IsEqualTo(typeof(WordprocessingDocument));
    }

    [Test]
    public async Task OXD021_CreateSpreadsheetDocument_ReturnsSpreadsheet()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument();
        await Assert.That(memDoc.GetDocumentType()).IsEqualTo(typeof(SpreadsheetDocument));
    }

    [Test]
    public async Task OXD022_CreatePresentationDocument_ReturnsPresentation()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
        await Assert.That(memDoc.GetDocumentType()).IsEqualTo(typeof(PresentationDocument));
    }

    // ── OpenXmlMemoryStreamDocument typed accessors ─────────────────────────

    [Test]
    public async Task OXD023_GetWordprocessingDocument_ReturnsOpenDocument()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
        using var wDoc = memDoc.GetWordprocessingDocument();
        await Assert.That(wDoc).IsNotNull();
        await Assert.That(wDoc.MainDocumentPart).IsNotNull();
    }

    [Test]
    public async Task OXD024_GetSpreadsheetDocument_ReturnsOpenDocument()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument();
        using var sDoc = memDoc.GetSpreadsheetDocument();
        await Assert.That(sDoc).IsNotNull();
        await Assert.That(sDoc.WorkbookPart).IsNotNull();
    }

    [Test]
    public async Task OXD025_GetPresentationDocument_ReturnsOpenDocument()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
        using var pDoc = memDoc.GetPresentationDocument();
        await Assert.That(pDoc).IsNotNull();
        await Assert.That(pDoc.PresentationPart).IsNotNull();
    }

    [Test]
    public async Task OXD026_GetWordprocessingDocument_WrongType_Throws()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument();
        await Assert.That(() => memDoc.GetWordprocessingDocument()).Throws<PowerToolsDocumentException>();
    }

    [Test]
    public async Task OXD027_GetSpreadsheetDocument_WrongType_Throws()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
        await Assert.That(() => memDoc.GetSpreadsheetDocument()).Throws<PowerToolsDocumentException>();
    }

    [Test]
    public async Task OXD028_GetPresentationDocument_WrongType_Throws()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
        await Assert.That(() => memDoc.GetPresentationDocument()).Throws<PowerToolsDocumentException>();
    }

    // ── OpenXmlMemoryStreamDocument.GetModified* round-trip ────────────────

    [Test]
    public async Task OXD030_GetModifiedWmlDocument_ProducesValidWmlDocument()
    {
        var source = new WmlDocument(DocxPath);
        using var memDoc = new OpenXmlMemoryStreamDocument(source);
        var modified = memDoc.GetModifiedWmlDocument();
        await Assert.That(modified).IsNotNull();
        await Assert.That(modified.GetDocumentType()).IsEqualTo(typeof(WordprocessingDocument));
    }

    [Test]
    public async Task OXD031_GetModifiedSmlDocument_ProducesValidSmlDocument()
    {
        var source = new SmlDocument(XlsxPath);
        using var memDoc = new OpenXmlMemoryStreamDocument(source);
        var modified = memDoc.GetModifiedSmlDocument();
        await Assert.That(modified).IsNotNull();
        await Assert.That(modified.GetDocumentType()).IsEqualTo(typeof(SpreadsheetDocument));
    }

    [Test]
    public async Task OXD032_GetModifiedPmlDocument_ProducesValidPmlDocument()
    {
        var source = new PmlDocument(PptxPath);
        using var memDoc = new OpenXmlMemoryStreamDocument(source);
        var modified = memDoc.GetModifiedPmlDocument();
        await Assert.That(modified).IsNotNull();
        await Assert.That(modified.GetDocumentType()).IsEqualTo(typeof(PresentationDocument));
    }

    // ── FromFileName error handling ─────────────────────────────────────────

    [Test]
    public async Task OXD040_FromFileName_NonOpenXmlFile_Throws()
    {
        // Write a plain-text file that is not a valid Open XML package
        var txtPath = Path.Combine(TempDir, "OXD040-not-openxml.txt");
        File.WriteAllText(txtPath, "This is not an Open XML document.");
        await Assert.That(() => OpenXmlPowerToolsDocument.FromFileName(txtPath)).Throws<PowerToolsDocumentException>();
    }

    // ── byte[] constructor ───────────────────────────────────────────────────

    [Test]
    public async Task OXD041_ByteArrayConstructor_CopiesBytesAndKeepsFileName()
    {
        var sourceBytes = File.ReadAllBytes(DocxPath);
        var doc = new WmlDocument("custom-name.docx", sourceBytes);
        await Assert.That(doc.DocumentByteArray).IsEquivalentTo(sourceBytes);
        await Assert.That(doc.FileName).IsEqualTo("custom-name.docx");

        // Verify the byte array is an independent copy, not a shared reference
        sourceBytes[0] = unchecked((byte)~sourceBytes[0]);
        await Assert.That(doc.DocumentByteArray[0]).IsNotEqualTo(sourceBytes[0]);
    }

    // ── Save ──────────────────────────────────────────────────────────────

    [Test]
    public async Task OXD042_Save_WithFileName_OverwritesOriginalFile()
    {
        var srcPath = Path.Combine(TempDir, "OXD042-source.docx");
        File.Copy(DocxPath, srcPath, overwrite: true);
        var doc = new WmlDocument(srcPath);
        doc.Save();
        var savedBytes = File.ReadAllBytes(srcPath);
        await Assert.That(savedBytes).IsEquivalentTo(doc.DocumentByteArray);
    }

    [Test]
    public async Task OXD043_Save_WithoutFileName_ThrowsInvalidOperationException()
    {
        var doc = new WmlDocument(null, File.ReadAllBytes(DocxPath));
        await Assert.That(() => doc.Save()).Throws<InvalidOperationException>();
    }

    // ── convertToTransitional constructors ───────────────────────────────────

    [Test]
    public async Task OXD044_FileNameConstructor_ConvertToTransitionalTrue_ProducesValidWmlDocument()
    {
        var doc = new WmlDocument(DocxPath, convertToTransitional: true);
        await Assert.That(doc.GetDocumentType()).IsEqualTo(typeof(WordprocessingDocument));
        await Assert.That(doc.FileName).IsEqualTo(DocxPath);
    }

    [Test]
    public async Task OXD045_ByteArrayConstructor_ConvertToTransitionalTrue_ProducesValidWmlDocument()
    {
        var doc = new WmlDocument(null, File.ReadAllBytes(DocxPath), convertToTransitional: true);
        await Assert.That(doc.GetDocumentType()).IsEqualTo(typeof(WordprocessingDocument));
    }

    [Test]
    public async Task OXD046_CopyConstructor_ConvertToTransitionalFalse_CopiesByteArray()
    {
        var original = new WmlDocument(DocxPath);
        var copy = new WmlDocument(original, convertToTransitional: false);
        await Assert.That(copy.DocumentByteArray).IsEquivalentTo(original.DocumentByteArray);
        await Assert.That(copy.FileName).IsEqualTo(original.FileName);
    }

    // ── SavePartAs ────────────────────────────────────────────────────────────

    [Test]
    public async Task OXD047_SavePartAs_WritesPartContentToFile()
    {
        var doc = new WmlDocument(DocxPath);
        using var memDoc = new OpenXmlMemoryStreamDocument(doc);
        using var wDoc = memDoc.GetWordprocessingDocument();
        var destPath = Path.Combine(TempDir, "OXD047-part.xml");
        OpenXmlPowerToolsDocument.SavePartAs(wDoc.MainDocumentPart, destPath);
        await Assert.That(File.Exists(destPath)).IsTrue();
        await Assert.That(new FileInfo(destPath).Length).IsGreaterThan(0);
    }

    // ── OpenXmlMemoryStreamDocument.CreatePackage / GetPackage ───────────────

    [Test]
    public async Task OXD048_CreatePackage_ReturnsEmptyPackage()
    {
        using var memDoc = OpenXmlMemoryStreamDocument.CreatePackage();
        var package = memDoc.GetPackage();
        await Assert.That(package).IsNotNull();
        await Assert.That(package.GetParts()).IsEmpty();
    }
}
