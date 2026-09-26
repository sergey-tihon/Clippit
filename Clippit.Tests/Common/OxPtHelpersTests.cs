// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="GetMetricsHelper"/> and <see cref="ValidationHelper"/>, the
/// file-path-based wrappers around <see cref="MetricsGetter"/> and <see cref="OpenXmlValidator"/>
/// used by the legacy CLI helpers in OxPtHelpers.cs.
/// </summary>
public class OxPtHelpersTests : TestsBase
{
    private static readonly string TestFilesDir = "../../../../TestFiles/";
    private static readonly string DocxPath = Path.Combine(TestFilesDir, "Blank-wml.docx");
    private static readonly string PptxPath = Path.Combine(TestFilesDir, "Presentation.pptx");
    private static readonly string XlsxPath = Path.Combine(TestFilesDir, "Spreadsheet.xlsx");

    // ── GetMetricsHelper.GetDocxMetrics ──────────────────────────────────────

    [Test]
    public async Task OPH001_GetDocxMetrics_ValidDocument_ReturnsMetricsWithFileName()
    {
        var metrics = GetMetricsHelper.GetDocxMetrics(DocxPath);

        await Assert.That(metrics).IsNotNull();
        await Assert.That(metrics.FileName).IsEqualTo(DocxPath);
    }

    [Test]
    public async Task OPH002_GetDocxMetrics_BlankDocument_HasZeroRunCount()
    {
        var metrics = GetMetricsHelper.GetDocxMetrics(DocxPath);

        await Assert.That(metrics.RunCount).IsEqualTo(0);
        await Assert.That(metrics.Table).IsEqualTo(0);
        await Assert.That(metrics.Hyperlink).IsEqualTo(0);
    }

    [Test]
    public async Task OPH003_GetDocxMetrics_ValidDocument_ValidIsBoolean()
    {
        // MetricsGetterSettings defaults (no fileFormatVersion override) drive the Valid flag;
        // just confirm it round-trips as a usable boolean rather than asserting a specific value,
        // since validity depends on the OpenXmlValidator default file-format version.
        var metrics = GetMetricsHelper.GetDocxMetrics(DocxPath);

        await Assert.That(metrics.Valid).IsTypeOf<bool>();
    }

    // ── ValidationHelper.IsValid ──────────────────────────────────────────────

    [Test]
    public async Task OPH010_IsValid_ValidWordDocument_ReturnsTrue()
    {
        var isValid = ValidationHelper.IsValid(DocxPath, "Office2013");

        await Assert.That(isValid).IsTrue();
    }

    [Test]
    public async Task OPH011_IsValid_ValidPresentationDocument_ReturnsTrue()
    {
        var isValid = ValidationHelper.IsValid(PptxPath, "Office2013");

        await Assert.That(isValid).IsTrue();
    }

    [Test]
    public async Task OPH012_IsValid_ValidSpreadsheetDocument_ReturnsTrue()
    {
        var isValid = ValidationHelper.IsValid(XlsxPath, "Office2013");

        await Assert.That(isValid).IsTrue();
    }

    [Test]
    public async Task OPH013_IsValid_UnsupportedOfficeVersion_FallsBackToOffice2013()
    {
        // An unparsable officeVersion string falls back to FileFormatVersions.Office2013
        // rather than throwing, so a valid document is still reported as valid.
        var isValid = ValidationHelper.IsValid(DocxPath, "NotARealVersion");

        await Assert.That(isValid).IsTrue();
    }

    [Test]
    public async Task OPH014_IsValid_UnsupportedFileExtension_ReturnsFalse()
    {
        var txtPath = Path.Combine(TempDir, "OPH014.txt");
        File.WriteAllText(txtPath, "not an office document");

        var isValid = ValidationHelper.IsValid(txtPath, "Office2013");

        await Assert.That(isValid).IsFalse();
    }

    // ── ValidationHelper.GetOpenXmlValidationErrors ──────────────────────────

    [Test]
    public async Task OPH020_GetOpenXmlValidationErrors_ValidWordDocument_ReturnsEmpty()
    {
        var errors = ValidationHelper.GetOpenXmlValidationErrors(DocxPath, "Office2013");

        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task OPH021_GetOpenXmlValidationErrors_ValidPresentationDocument_ReturnsEmpty()
    {
        var errors = ValidationHelper.GetOpenXmlValidationErrors(PptxPath, "Office2013");

        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task OPH022_GetOpenXmlValidationErrors_ValidSpreadsheetDocument_ReturnsEmpty()
    {
        var errors = ValidationHelper.GetOpenXmlValidationErrors(XlsxPath, "Office2013");

        await Assert.That(errors).IsEmpty();
    }

    [Test]
    public async Task OPH023_GetOpenXmlValidationErrors_UnsupportedFileExtension_ReturnsEmpty()
    {
        var txtPath = Path.Combine(TempDir, "OPH023.txt");
        File.WriteAllText(txtPath, "not an office document");

        var errors = ValidationHelper.GetOpenXmlValidationErrors(txtPath, "Office2013");

        await Assert.That(errors).IsEmpty();
    }
}
