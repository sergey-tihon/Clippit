// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO.Packaging;
using System.Xml.Linq;

namespace Clippit.Tests.Common;

/// <summary>
/// Tests for <see cref="FlatOpc"/> binary-part round-trip through Flat OPC format.
/// </summary>
public class FlatOpcTests
{
    private static readonly XNamespace s_pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
    private static readonly XNamespace s_rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    private static XDocument BuildFlatOpcWithBinaryPart(string base64Data)
    {
        return new XDocument(
            new XElement(
                s_pkg + "package",
                new XAttribute(XNamespace.Xmlns + "pkg", s_pkg.ToString()),
                new XElement(
                    s_pkg + "part",
                    new XAttribute(s_pkg + "name", "/_rels/.rels"),
                    new XAttribute(s_pkg + "contentType", "application/vnd.openxmlformats-package.relationships+xml"),
                    new XElement(s_pkg + "xmlData", new XElement(s_rel + "Relationships"))
                ),
                new XElement(
                    s_pkg + "part",
                    new XAttribute(s_pkg + "name", "/media/data.bin"),
                    new XAttribute(s_pkg + "contentType", "application/octet-stream"),
                    new XAttribute(s_pkg + "compression", "store"),
                    new XElement(s_pkg + "binaryData", base64Data)
                )
            )
        );
    }

    private static byte[] ReadBinaryPartFromPackage(string packagePath)
    {
        using var package = Package.Open(packagePath, FileMode.Open, FileAccess.Read);
        var part = package.GetPart(new Uri("/media/data.bin", UriKind.Relative));
        using var stream = part.GetStream();
        using var ms = new MemoryStream();
        stream.CopyTo(ms);
        return ms.ToArray();
    }

    [Test]
    [Arguments("\n")]
    [Arguments("\r\n")]
    public async Task FOC_001_FlatToOpc_BinaryPartWithChunkedBase64_RoundTrips(string lineEnding)
    {
        var original = Enumerable.Range(0, 256).Select(i => (byte)i).ToArray();
        var flat = Convert.ToBase64String(original);
        var chunked = Base64.ChunkBase64(flat, appendTrailingNewline: true).ReplaceLineEndings(lineEnding);

        var doc = BuildFlatOpcWithBinaryPart(chunked);

        var outputPath = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".zip");
        try
        {
            FlatOpc.FlatToOpc(doc, outputPath);
            var recovered = ReadBinaryPartFromPackage(outputPath);
            await Assert.That(recovered).IsEquivalentTo(original);
        }
        finally
        {
            if (File.Exists(outputPath))
                File.Delete(outputPath);
        }
    }
}
