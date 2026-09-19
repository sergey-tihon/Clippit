// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace Clippit.Tests.PowerPoint;

public class PtUtilTests
{
    [Test]
    [Arguments("PU/PU001-Test001.mht")]
    public async Task PU001(string name)
    {
        var sourceDir = new DirectoryInfo("../../../../TestFiles/");
        var sourceMht = new FileInfo(Path.Combine(sourceDir.FullName, name));
        var src = await File.ReadAllTextAsync(sourceMht.FullName);
        var p = MhtParser.Parse(src);

        await Assert.That(p.ContentType).IsNotNull();
        await Assert.That(p.MimeVersion).IsNotNull();
        await Assert.That(p.Parts).IsNotEmpty();
        await Assert.That(p.Parts).DoesNotContain(part => part.ContentType == null || part.ContentLocation == null);
    }

    [Test]
    public async Task PU002_MixedCaseBoundaryAndCharSetAreParsed()
    {
        var newLine = Environment.NewLine;
        var src = string.Join(
            newLine,
            "MIME-Version: 1.0",
            "Content-Type: multipart/related; Boundary=\"----=_NextPart_Test\"",
            "",
            "------=_NextPart_Test",
            "Content-Location: file:///C:/Test.htm",
            "Content-Transfer-Encoding: quoted-printable",
            "Content-Type: text/html; Charset=\"utf-8\"",
            "",
            "<html></html>",
            "------=_NextPart_Test--",
            ""
        );

        var p = MhtParser.Parse(src);

        await Assert.That(p.ContentType).IsEqualTo("multipart/related");
        var part = p.Parts.Single();
        await Assert.That(part.ContentType).IsEqualTo("text/html");
        await Assert.That(part.CharSet).IsEqualTo("utf-8");
    }
}
