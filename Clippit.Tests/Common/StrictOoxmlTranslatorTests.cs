// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Internal;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="StrictOoxmlTranslator"/>.
/// Verifies that strict ISO 29500 namespace URIs and relationship types are
/// correctly mapped to their Transitional OOXML equivalents.
/// </summary>
public class StrictOoxmlTranslatorTests
{
    // ── TranslateNamespace: known mappings ──────────────────────────────────

    [Test]
    public async Task SOT001_TranslateNamespace_WordprocessingMlMain_MapsToTransitional()
    {
        var result = StrictOoxmlTranslator.TranslateNamespace("http://purl.oclc.org/ooxml/wordprocessingml/main");
        await Assert.That(result).IsEqualTo("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    }

    [Test]
    public async Task SOT002_TranslateNamespace_PresentationMlMain_MapsToTransitional()
    {
        var result = StrictOoxmlTranslator.TranslateNamespace("http://purl.oclc.org/ooxml/presentationml/main");
        await Assert.That(result).IsEqualTo("http://schemas.openxmlformats.org/presentationml/2006/main");
    }

    [Test]
    public async Task SOT003_TranslateNamespace_SpreadsheetMlMain_MapsToTransitional()
    {
        var result = StrictOoxmlTranslator.TranslateNamespace("http://purl.oclc.org/ooxml/spreadsheetml/main");
        await Assert.That(result).IsEqualTo("http://schemas.openxmlformats.org/spreadsheetml/2006/main");
    }

    [Test]
    public async Task SOT004_TranslateNamespace_DrawingMlMain_MapsToTransitional()
    {
        var result = StrictOoxmlTranslator.TranslateNamespace("http://purl.oclc.org/ooxml/drawingml/main");
        await Assert.That(result).IsEqualTo("http://schemas.openxmlformats.org/drawingml/2006/main");
    }

    // ── TranslateNamespace: spec anomalies (hyphens) ────────────────────────

    [Test]
    public async Task SOT005_TranslateNamespace_CustomProperties_MapsToHyphenatedForm()
    {
        // "customProperties" (no hyphen) → "custom-properties" (hyphen) per spec
        var result = StrictOoxmlTranslator.TranslateNamespace(
            "http://purl.oclc.org/ooxml/officeDocument/customProperties"
        );
        await Assert.That(result).IsEqualTo("http://schemas.openxmlformats.org/officeDocument/2006/custom-properties");
    }

    [Test]
    public async Task SOT006_TranslateNamespace_ExtendedProperties_MapsToHyphenatedForm()
    {
        // "extendedProperties" (no hyphen) → "extended-properties" (hyphen) per spec
        var result = StrictOoxmlTranslator.TranslateNamespace(
            "http://purl.oclc.org/ooxml/officeDocument/extendedProperties"
        );
        await Assert
            .That(result)
            .IsEqualTo("http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
    }

    // ── TranslateNamespace: unknown/already-transitional URIs ───────────────

    [Test]
    public async Task SOT007_TranslateNamespace_AlreadyTransitionalUri_ReturnedUnchanged()
    {
        const string transitional = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var result = StrictOoxmlTranslator.TranslateNamespace(transitional);
        await Assert.That(result).IsEqualTo(transitional);
    }

    [Test]
    public async Task SOT008_TranslateNamespace_ArbitraryUnknownUri_ReturnedUnchanged()
    {
        const string unknown = "http://example.com/some/custom/namespace";
        var result = StrictOoxmlTranslator.TranslateNamespace(unknown);
        await Assert.That(result).IsEqualTo(unknown);
    }

    [Test]
    public async Task SOT009_TranslateNamespace_EmptyString_ReturnedUnchanged()
    {
        var result = StrictOoxmlTranslator.TranslateNamespace(string.Empty);
        await Assert.That(result).IsEqualTo(string.Empty);
    }

    // ── TranslateNamespace: case-sensitivity (Ordinal comparer) ─────────────

    [Test]
    public async Task SOT010_TranslateNamespace_WrongCase_ReturnedUnchanged()
    {
        // The lookup uses StringComparer.Ordinal; upper-cased input must not match.
        const string wrongCase = "http://PURL.OCLC.ORG/ooxml/wordprocessingml/main";
        var result = StrictOoxmlTranslator.TranslateNamespace(wrongCase);
        await Assert.That(result).IsEqualTo(wrongCase);
    }

    // ── TranslateRelationshipType: known mappings ───────────────────────────

    [Test]
    public async Task SOT011_TranslateRelationshipType_SlideRel_MapsToTransitional()
    {
        var result = StrictOoxmlTranslator.TranslateRelationshipType(
            "http://purl.oclc.org/ooxml/officeDocument/relationships/slide"
        );
        await Assert
            .That(result)
            .IsEqualTo("http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide");
    }

    [Test]
    public async Task SOT012_TranslateRelationshipType_StylesRel_MapsToTransitional()
    {
        var result = StrictOoxmlTranslator.TranslateRelationshipType(
            "http://purl.oclc.org/ooxml/officeDocument/relationships/styles"
        );
        await Assert
            .That(result)
            .IsEqualTo("http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles");
    }

    [Test]
    public async Task SOT013_TranslateRelationshipType_ImageRel_MapsToTransitional()
    {
        var result = StrictOoxmlTranslator.TranslateRelationshipType(
            "http://purl.oclc.org/ooxml/officeDocument/relationships/image"
        );
        await Assert
            .That(result)
            .IsEqualTo("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image");
    }

    // ── TranslateRelationshipType: unknown/already-transitional URIs ─────────

    [Test]
    public async Task SOT014_TranslateRelationshipType_AlreadyTransitional_ReturnedUnchanged()
    {
        const string transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
        var result = StrictOoxmlTranslator.TranslateRelationshipType(transitional);
        await Assert.That(result).IsEqualTo(transitional);
    }

    [Test]
    public async Task SOT015_TranslateRelationshipType_ArbitraryUnknownUri_ReturnedUnchanged()
    {
        const string unknown = "http://example.com/custom/relationship";
        var result = StrictOoxmlTranslator.TranslateRelationshipType(unknown);
        await Assert.That(result).IsEqualTo(unknown);
    }

    // ── TranslateNamespace: dublin-core identity mappings ───────────────────

    [Test]
    public async Task SOT016_TranslateNamespace_DublinCoreElements_ReturnsSameUri()
    {
        // The identity mappings in the table are present for completeness; they must
        // be returned as-is (the value equals the key).
        const string dc = "http://purl.org/dc/elements/1.1/";
        var result = StrictOoxmlTranslator.TranslateNamespace(dc);
        await Assert.That(result).IsEqualTo(dc);
    }
}
