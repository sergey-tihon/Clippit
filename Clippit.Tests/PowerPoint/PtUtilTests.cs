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
}
