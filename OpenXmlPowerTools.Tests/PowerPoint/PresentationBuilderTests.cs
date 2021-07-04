// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using Clippit.PowerPoint;
using DocumentFormat.OpenXml.Packaging;
using Xunit;
using Xunit.Abstractions;

#if !ELIDE_XUNIT_TESTS

namespace Clippit.Tests.PowerPoint
{
    public class PresentationBuilderTests : TestsBase
    {
        public PresentationBuilderTests(ITestOutputHelper log) : base(log)
        {
        }
        
        [Fact]
        public void PB001_Formatting()
        {
            var name1 = "PB001-Input1.pptx";
            var name2 = "PB001-Input2.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));

            var sources = new List<SlideSource>
            {
                new(new PmlDocument(source1Pptx.FullName), 1, true),
                new(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB001-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        }

        [Fact]
        public void PB002_Formatting()
        {
            var name2 = "PB001-Input2.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));

            var sources = new List<SlideSource>
            {
                new(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB002-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        }

        [Fact]
        public void PB003_Formatting()
        {
            var name1 = "PB001-Input1.pptx";
            var name2 = "PB001-Input3.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));

            var sources = new List<SlideSource>
            {
                new(new PmlDocument(source1Pptx.FullName), 1, true),
                new(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB003-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        }

        [Fact]
        public void PB004_Formatting()
        {
            var name1 = "PB001-Input1.pptx";
            var name2 = "PB001-Input3.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));

            var sources = new List<SlideSource>
            {
                new(new PmlDocument(source2Pptx.FullName), 0, true),
                new(new PmlDocument(source1Pptx.FullName), 1, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB004-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        }

        [Fact]
        public void PB005_Formatting()
        {
            var name1 = "PB001-Input1.pptx";
            var name2 = "PB001-Input3.pptx";
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var source1Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name1));
            var source2Pptx = new FileInfo(Path.Combine(sourceDir.FullName, name2));

            var sources = new List<SlideSource>
            {
                new(new PmlDocument(source2Pptx.FullName), 0, 0, true),
                new(new PmlDocument(source1Pptx.FullName), 1, true),
                new(new PmlDocument(source2Pptx.FullName), 0, true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB005-Formatting.pptx"));
            PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);
        }

        [Fact()]
        public void PB006_VideoFormats()
        {
            // This presentation contains videos with content types video/mp4, video/quicktime, video/unknown, video/x-ms-asf, and video/x-msvideo.
            var sourceDir = new DirectoryInfo("../../../../TestFiles/");
            var sourcePptx = new FileInfo(Path.Combine(sourceDir.FullName, "PP006-Videos.pptx"));

            var oldMediaDataContentTypes = GetMediaDataContentTypes(sourcePptx);

            var sources = new List<SlideSource>
            {
                new(new PmlDocument(sourcePptx.FullName), true),
            };
            var processedDestPptx = new FileInfo(Path.Combine(TempDir, "PB006-Videos.pptx"));
            PresentationBuilder.BuildPresentation(sources).SaveAs(processedDestPptx.FullName);

            var newMediaDataContentTypes = GetMediaDataContentTypes(processedDestPptx);

            Assert.Equal(oldMediaDataContentTypes, newMediaDataContentTypes);
        }

        private static string[] GetMediaDataContentTypes(FileInfo fi)
        {
            using var ptDoc = PresentationDocument.Open(fi.FullName, false);
            return ptDoc.PresentationPart.SlideParts.SelectMany(
                    p => p.DataPartReferenceRelationships.Select(d => d.DataPart.ContentType))
                .Distinct()
                .OrderBy(m => m)
                .ToArray();
        }
    }
}

#endif
