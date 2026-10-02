// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml;
using System.Xml.Linq;
using Clippit.Internal;

namespace Clippit.Tests.Common;

/// <summary>
/// Unit tests for <see cref="StrictTranslatingXmlReader"/>.
/// Verifies that Strict ISO/IEC 29500 namespace URIs are translated to their
/// Transitional equivalents while reading, independent of any package/part machinery.
/// </summary>
public class StrictTranslatingXmlReaderTests
{
    private static XDocument LoadTranslated(string xml)
    {
        using var stringReader = new StringReader(xml);
        using var xmlReader = XmlReader.Create(stringReader);
        using var translatingReader = new StrictTranslatingXmlReader(xmlReader);
        return XDocument.Load(translatingReader);
    }

    [Test]
    public async Task STXR001_ElementNamespace_IsTranslatedToTransitional()
    {
        const string xml = """<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" />""";

        var doc = LoadTranslated(xml);

        await Assert
            .That(doc.Root!.Name.NamespaceName)
            .IsEqualTo("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    }

    [Test]
    public async Task STXR002_XmlnsDeclarationValue_IsTranslatedToTransitional()
    {
        const string xml = """<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" />""";

        var doc = LoadTranslated(xml);

        // The xmlns:w declaration attribute value must also be translated so the
        // serialized document stays internally consistent with the element namespace.
        var xmlnsAttr = doc.Root!.Attribute(XNamespace.Xmlns + "w");
        await Assert.That(xmlnsAttr).IsNotNull();
        await Assert.That(xmlnsAttr!.Value).IsEqualTo("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    }

    [Test]
    public async Task STXR003_GraphicDataUriAttribute_IsTranslatedToTransitional()
    {
        const string xml = """
            <a:graphicFrame xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">
              <a:graphic>
                <a:graphicData uri="http://purl.oclc.org/ooxml/drawingml/chart" />
              </a:graphic>
            </a:graphicFrame>
            """;

        var doc = LoadTranslated(xml);

        var graphicData = doc.Descendants().Single(e => e.Name.LocalName == "graphicData");
        await Assert
            .That(graphicData.Attribute("uri")!.Value)
            .IsEqualTo("http://schemas.openxmlformats.org/drawingml/2006/chart");
    }

    [Test]
    public async Task STXR004_RegularAttributeValue_IsNotTranslated()
    {
        const string xml = """
            <w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main">
              <w:body w:val="http://purl.oclc.org/ooxml/wordprocessingml/main" />
            </w:document>
            """;

        var doc = LoadTranslated(xml);

        // Only xmlns:* declarations and uri= attributes get translated - any other
        // attribute value that happens to look like a Strict namespace must pass through unchanged.
        var body = doc.Descendants().Single(e => e.Name.LocalName == "body");
        await Assert
            .That(body.Attribute(doc.Root!.Name.Namespace + "val")!.Value)
            .IsEqualTo("http://purl.oclc.org/ooxml/wordprocessingml/main");
    }

    [Test]
    public async Task STXR005_UnknownNamespace_PassesThroughUnchanged()
    {
        const string xml = """<root xmlns="http://example.com/not-strict" />""";

        var doc = LoadTranslated(xml);

        await Assert.That(doc.Root!.Name.NamespaceName).IsEqualTo("http://example.com/not-strict");
    }

    [Test]
    public async Task STXR006_LookupNamespace_ReturnsTranslatedUri()
    {
        const string xml = """<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" />""";
        using var stringReader = new StringReader(xml);
        using var xmlReader = XmlReader.Create(stringReader);
        using var translatingReader = new StrictTranslatingXmlReader(xmlReader);

        translatingReader.Read();

        await Assert
            .That(translatingReader.LookupNamespace("w"))
            .IsEqualTo("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
    }
}
