// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using Clippit;

namespace Clippit.Tests.Common;

public class UriFixerTests
{
    private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

    // Build a minimal in-memory ZIP that contains one .rels entry.
    private static MemoryStream BuildZipWithRels(string relsXml)
    {
        var ms = new MemoryStream();
        using (var za = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = za.CreateEntry("_rels/.rels");
            using var writer = new StreamWriter(entry.Open());
            writer.Write(relsXml);
        }
        ms.Position = 0;
        return ms;
    }

    private static XDocument ReadRelsFromZip(MemoryStream ms)
    {
        ms.Position = 0;
        using var za = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: true);
        var entry = za.GetEntry("_rels/.rels")!;
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    // ── UF001: valid URI is left unchanged ──────────────────────────────────

    [Test]
    public async Task UF001_ValidUri_IsNotModified()
    {
        const string validUrl = "https://example.com/page";
        var relsXml = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="{RelNs}">
              <Relationship Id="rId1" Type="http://example.com/rel" TargetMode="External" Target="{validUrl}" />
            </Relationships>
            """;

        var ms = BuildZipWithRels(relsXml);
        UriFixer.FixInvalidUri(ms, leaveOpen: true);

        var xdoc = ReadRelsFromZip(ms);
        var target = xdoc.Descendants(RelNs + "Relationship").Single().Attribute("Target")?.Value;
        await Assert.That(target).IsEqualTo(validUrl);
    }

    // ── UF002: invalid URI is replaced with the default placeholder ─────────

    [Test]
    public async Task UF002_InvalidUri_IsReplacedWithPlaceholder()
    {
        const string invalidUrl = "not a valid uri: \\bad";
        var relsXml = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="{RelNs}">
              <Relationship Id="rId1" Type="http://example.com/rel" TargetMode="External" Target="{invalidUrl}" />
            </Relationships>
            """;

        var ms = BuildZipWithRels(relsXml);
        UriFixer.FixInvalidUri(ms, leaveOpen: true);

        var xdoc = ReadRelsFromZip(ms);
        var target = xdoc.Descendants(RelNs + "Relationship").Single().Attribute("Target")?.Value;
        await Assert.That(target).IsEqualTo("https://example.invalid");
    }

    // ── UF003: custom handler receives the invalid URI and its result is used ─

    [Test]
    public async Task UF003_CustomHandler_IsCalledWithInvalidUri()
    {
        const string invalidUrl = "bad uri here";
        string? capturedUri = null;
        var replacement = new Uri("https://replaced.example.com");

        var relsXml = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="{RelNs}">
              <Relationship Id="rId1" Type="http://example.com/rel" TargetMode="External" Target="{invalidUrl}" />
            </Relationships>
            """;

        var ms = BuildZipWithRels(relsXml);
        UriFixer.FixInvalidUri(
            ms,
            uri =>
            {
                capturedUri = uri;
                return replacement;
            },
            leaveOpen: true
        );

        await Assert.That(capturedUri).IsEqualTo(invalidUrl);
        var xdoc = ReadRelsFromZip(ms);
        var target = xdoc.Descendants(RelNs + "Relationship").Single().Attribute("Target")?.Value;
        await Assert.That(target).IsEqualTo(replacement.OriginalString);
    }

    // ── UF004: Internal relationships (no TargetMode=External) are not touched ─

    [Test]
    public async Task UF004_InternalRelationship_IsNotModified()
    {
        const string internalTarget = "word/document.xml";
        var relsXml = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="{RelNs}">
              <Relationship Id="rId1" Type="http://example.com/rel" Target="{internalTarget}" />
            </Relationships>
            """;

        var ms = BuildZipWithRels(relsXml);
        UriFixer.FixInvalidUri(ms, leaveOpen: true);

        var xdoc = ReadRelsFromZip(ms);
        var target = xdoc.Descendants(RelNs + "Relationship").Single().Attribute("Target")?.Value;
        await Assert.That(target).IsEqualTo(internalTarget);
    }

    // ── UF005: leaveOpen=true keeps the stream usable ───────────────────────

    [Test]
    public async Task UF005_LeaveOpenTrue_StreamRemainsUsable()
    {
        var relsXml = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="{RelNs}" />
            """;

        var ms = BuildZipWithRels(relsXml);
        UriFixer.FixInvalidUri(ms, leaveOpen: true);

        // Stream must still be readable after the call
        await Assert.That(ms.CanRead).IsTrue();
        ms.Position = 0;
        await Assert.That(ms.Length).IsGreaterThan(0L);
    }

    // ── UF006: multiple URIs in the same .rels entry are all fixed ───────────

    [Test]
    public async Task UF006_MultipleInvalidUris_AllReplaced()
    {
        var relsXml = $"""
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="{RelNs}">
              <Relationship Id="rId1" Type="http://example.com/rel" TargetMode="External" Target="bad uri 1" />
              <Relationship Id="rId2" Type="http://example.com/rel" TargetMode="External" Target="bad uri 2" />
            </Relationships>
            """;

        var ms = BuildZipWithRels(relsXml);
        UriFixer.FixInvalidUri(ms, leaveOpen: true);

        var xdoc = ReadRelsFromZip(ms);
        var targets = xdoc.Descendants(RelNs + "Relationship").Select(r => r.Attribute("Target")?.Value).ToList();
        await Assert.That(targets).HasCount(2);
        await Assert.That(targets[0]).IsEqualTo("https://example.invalid");
        await Assert.That(targets[1]).IsEqualTo("https://example.invalid");
    }
}
