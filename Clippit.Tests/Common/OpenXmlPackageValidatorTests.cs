// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Core;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="OpenXmlPackageValidator.Validate"/>, which combines
/// OpenXml schema validation with relationship-reference validation and returns a
/// unified <see cref="OpenXmlValidationResult"/>.
/// </summary>
public class OpenXmlPackageValidatorTests : TestsBase
{
    private static readonly DirectoryInfo s_testFiles = new("../../../../TestFiles");

    // ── OPV001: a well-formed DOCX produces no diagnostics ──────────────────

    [Test]
    public async Task OPV001_CleanDocx_NoDiagnostics()
    {
        using var doc = WordprocessingDocument.Open(Path.Combine(s_testFiles.FullName, "Blank-wml.docx"), false);

        var result = OpenXmlPackageValidator.Validate(doc);

        await Assert.That(result.Valid).IsTrue();
        await Assert.That(result.Diagnostics).IsEmpty();
    }

    // ── OPV002: a well-formed PPTX produces no diagnostics ──────────────────

    [Test]
    public async Task OPV002_CleanPptx_NoDiagnostics()
    {
        using var pres = PresentationDocument.Open(Path.Combine(s_testFiles.FullName, "PB001-Input1.pptx"), false);

        var result = OpenXmlPackageValidator.Validate(pres);

        await Assert.That(result.Valid).IsTrue();
        await Assert.That(result.Diagnostics).IsEmpty();
    }

    // ── OPV003: null package throws ArgumentNullException ───────────────────

    [Test]
    public async Task OPV003_NullPackage_ThrowsArgumentNullException()
    {
#pragma warning disable CS8625 // Cannot convert null literal to non-nullable reference type.
        await Assert.That(() => OpenXmlPackageValidator.Validate(null)).Throws<ArgumentNullException>();
#pragma warning restore CS8625 // Cannot convert null literal to non-nullable reference type.
    }

    // ── OPV004: result reflects the OfficeVersion from options ──────────────

    [Test]
    public async Task OPV004_CustomOfficeVersion_ReflectedInResult()
    {
        using var doc = WordprocessingDocument.Open(Path.Combine(s_testFiles.FullName, "Blank-wml.docx"), false);
        var options = new OpenXmlValidationOptions { OfficeVersion = FileFormatVersions.Office2019 };

        var result = OpenXmlPackageValidator.Validate(doc, options);

        await Assert.That(result.OfficeVersion).IsEqualTo(FileFormatVersions.Office2019);
    }

    // ── OPV005: default options uses Microsoft365 version ───────────────────

    [Test]
    public async Task OPV005_DefaultOptions_UsesMicrosoft365Version()
    {
        using var doc = WordprocessingDocument.Open(Path.Combine(s_testFiles.FullName, "Blank-wml.docx"), false);

        var result = OpenXmlPackageValidator.Validate(doc);

        await Assert.That(result.OfficeVersion).IsEqualTo(FileFormatVersions.Microsoft365);
    }

    // ── OPV006: a DOCX with a dangling relationship id produces a diagnostic

    [Test]
    public async Task OPV006_DanglingRelationshipId_ProducesDiagnostic()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document, true))
        {
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new Document(
                new Body(
                    new Paragraph(
                        new Run(
                            new Drawing(
                                new DocumentFormat.OpenXml.Drawing.Wordprocessing.Inline(
                                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.Extent
                                    {
                                        Cx = 1000000L,
                                        Cy = 1000000L,
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.EffectExtent
                                    {
                                        LeftEdge = 0L,
                                        TopEdge = 0L,
                                        RightEdge = 0L,
                                        BottomEdge = 0L,
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties
                                    {
                                        Id = 1U,
                                        Name = "Image 1",
                                    },
                                    new DocumentFormat.OpenXml.Drawing.Graphic(
                                        new DocumentFormat.OpenXml.Drawing.GraphicData(
                                            new DocumentFormat.OpenXml.Drawing.Pictures.Picture(
                                                new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureProperties(
                                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualDrawingProperties
                                                    {
                                                        Id = 1U,
                                                        Name = "img.png",
                                                    },
                                                    new DocumentFormat.OpenXml.Drawing.Pictures.NonVisualPictureDrawingProperties()
                                                ),
                                                new DocumentFormat.OpenXml.Drawing.Pictures.BlipFill(
                                                    new DocumentFormat.OpenXml.Drawing.Blip { Embed = "rId999" },
                                                    new DocumentFormat.OpenXml.Drawing.Stretch(
                                                        new DocumentFormat.OpenXml.Drawing.FillRectangle()
                                                    )
                                                ),
                                                new DocumentFormat.OpenXml.Drawing.Pictures.ShapeProperties(
                                                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                                                        new DocumentFormat.OpenXml.Drawing.Offset { X = 0L, Y = 0L },
                                                        new DocumentFormat.OpenXml.Drawing.Extents
                                                        {
                                                            Cx = 1000000L,
                                                            Cy = 1000000L,
                                                        }
                                                    ),
                                                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                                                        new DocumentFormat.OpenXml.Drawing.AdjustValueList()
                                                    )
                                                    {
                                                        Preset = DocumentFormat
                                                            .OpenXml
                                                            .Drawing
                                                            .ShapeTypeValues
                                                            .Rectangle,
                                                    }
                                                )
                                            )
                                        )
                                        {
                                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture",
                                        }
                                    )
                                )
                                {
                                    DistanceFromTop = 0U,
                                    DistanceFromBottom = 0U,
                                    DistanceFromLeft = 0U,
                                    DistanceFromRight = 0U,
                                }
                            )
                        )
                    )
                )
            );
            doc.MainDocumentPart!.Document.Save();
        }

        ms.Position = 0;
        using var docToValidate = WordprocessingDocument.Open(ms, false);
        var result = OpenXmlPackageValidator.Validate(docToValidate);

        await Assert.That(result.Valid).IsFalse();
        var danglingDiag = result.Diagnostics.FirstOrDefault(d =>
            d.Kind == OpenXmlValidationDiagnosticKinds.Relationship
        );
        await Assert.That(danglingDiag).IsNotNull();
        await Assert.That(danglingDiag!.RelationshipId).IsEqualTo("rId999");
    }

    // ── OPV007: Valid property false when diagnostics exist ──────────────────

    [Test]
    public async Task OPV007_ValidPropertyFalse_WhenDiagnosticsExist()
    {
        var result = new OpenXmlValidationResult
        {
            OfficeVersion = FileFormatVersions.Microsoft365,
            Diagnostics =
            [
                new OpenXmlValidationDiagnostic
                {
                    Kind = OpenXmlValidationDiagnosticKinds.Schema,
                    Description = "test error",
                },
            ],
        };

        await Assert.That(result.Valid).IsFalse();
        await Assert.That(result.Diagnostics).HasCount(1);
    }

    // ── OPV008: Valid property true when diagnostics empty ───────────────────

    [Test]
    public async Task OPV008_ValidPropertyTrue_WhenNoDiagnostics()
    {
        var result = new OpenXmlValidationResult { OfficeVersion = FileFormatVersions.Microsoft365, Diagnostics = [] };

        await Assert.That(result.Valid).IsTrue();
    }
}
