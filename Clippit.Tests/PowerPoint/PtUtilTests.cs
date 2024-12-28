﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace Clippit.Tests.PowerPoint
{
    public class PtUtilTests
    {
        [Theory]
        [InlineData("PU/PU001-Test001.mht")]
        public void PU001(string name)
        {
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourceMht = new FileInfo(Path.Combine(sourceDir.FullName, name));
            var src = File.ReadAllText(sourceMht.FullName);
            var p = MhtParser.Parse(src);
            Assert.NotNull(p.ContentType);
            Assert.NotNull(p.MimeVersion);
            Assert.NotEmpty(p.Parts);
            Assert.DoesNotContain(p.Parts, part => part.ContentType == null || part.ContentLocation == null);
        }
    }
}

#endif
