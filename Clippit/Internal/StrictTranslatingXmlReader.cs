// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml;

namespace Clippit.Internal;

/// <summary>
/// An <see cref="XmlReader"/> wrapper that transparently translates ISO/IEC 29500 Strict
/// XML namespace URIs to their OOXML Transitional equivalents as the document is parsed.
/// </summary>
/// <remarks>
/// <para>
/// Strict OOXML documents (conformance="strict") use <c>http://purl.oclc.org/ooxml/…</c>
/// namespace URIs.  All Clippit code queries the XML tree using the Transitional
/// <c>http://schemas.openxmlformats.org/…</c> constants.  Wrapping the underlying
/// <see cref="XmlReader"/> with this class makes the translation happen at parse time,
/// so the resulting <see cref="System.Xml.Linq.XDocument"/> already carries Transitional
/// namespace URIs — no post-processing or whole-file byte-array conversion is required.
/// </para>
/// <para>
/// Two members are overridden:
/// <list type="bullet">
///   <item><see cref="NamespaceURI"/> — returns the translated URI for every element and
///   attribute node.</item>
///   <item><see cref="Value"/> — translates the value of <c>xmlns:*</c> declaration
///   attributes so that the namespace declarations stored in the XDocument remain
///   consistent with the element namespace URIs when the document is serialised back.</item>
/// </list>
/// <see cref="LookupNamespace"/> is also overridden so that any caller resolving a prefix
/// to a URI receives the Transitional form.
/// </para>
/// <para>
/// All other members delegate to the inner reader unchanged.
/// The inner reader is disposed when this instance is disposed.
/// </para>
/// </remarks>
internal sealed class StrictTranslatingXmlReader(XmlReader inner) : XmlReader
{
    private const string XmlnsNamespace = "http://www.w3.org/2000/xmlns/";

    // ── Abstract members — delegated to inner ────────────────────────────────

    public override int AttributeCount => inner.AttributeCount;
    public override string BaseURI => inner.BaseURI;
    public override int Depth => inner.Depth;
    public override bool EOF => inner.EOF;
    public override bool IsEmptyElement => inner.IsEmptyElement;
    public override string LocalName => inner.LocalName;
    public override XmlNameTable NameTable => inner.NameTable;
    public override XmlNodeType NodeType => inner.NodeType;
    public override string Prefix => inner.Prefix;
    public override ReadState ReadState => inner.ReadState;

    public override string? GetAttribute(int i) => inner.GetAttribute(i);

    public override string? GetAttribute(string name) => inner.GetAttribute(name);

    public override string? GetAttribute(string localName, string? namespaceURI) =>
        inner.GetAttribute(localName, namespaceURI);

    public override void MoveToAttribute(int i) => inner.MoveToAttribute(i);

    public override bool MoveToAttribute(string name) => inner.MoveToAttribute(name);

    public override bool MoveToAttribute(string localName, string? namespaceURI) =>
        inner.MoveToAttribute(localName, namespaceURI);

    public override bool MoveToElement() => inner.MoveToElement();

    public override bool MoveToFirstAttribute() => inner.MoveToFirstAttribute();

    public override bool MoveToNextAttribute() => inner.MoveToNextAttribute();

    public override bool Read() => inner.Read();

    public override bool ReadAttributeValue() => inner.ReadAttributeValue();

    public override void ResolveEntity() => inner.ResolveEntity();

    // ── Overridden members — translate Strict → Transitional ────────────────

    /// <summary>
    /// Returns the Transitional namespace URI for the current element or attribute node.
    /// </summary>
    public override string NamespaceURI => StrictOoxmlTranslator.TranslateNamespace(inner.NamespaceURI);

    /// <summary>
    /// Returns the node value, translating the URI for <c>xmlns:*</c> declaration attributes
    /// so that the namespace declarations in the loaded XDocument match the translated element
    /// namespace URIs and remain consistent on serialization.
    /// </summary>
    public override string Value =>
        inner.NodeType == XmlNodeType.Attribute && inner.NamespaceURI == XmlnsNamespace
            ? StrictOoxmlTranslator.TranslateNamespace(inner.Value)
            : inner.Value;

    /// <summary>
    /// Resolves a namespace prefix to its URI, returning the Transitional form when the
    /// prefix is bound to a Strict namespace URI.
    /// </summary>
    public override string? LookupNamespace(string prefix)
    {
        var uri = inner.LookupNamespace(prefix);
        return uri is null ? null : StrictOoxmlTranslator.TranslateNamespace(uri);
    }

    // ── Disposal ─────────────────────────────────────────────────────────────

    protected override void Dispose(bool disposing)
    {
        if (disposing)
            inner.Dispose();
        base.Dispose(disposing);
    }
}
