// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests
{
    /// <summary>
    /// Base class for unit tests providing utility methods.
    /// </summary>
    public class TestsBase
    {
        protected static void CreateEmptyWordprocessingDocument(Stream stream)
        {
            using var wordDocument = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document);
            var part = wordDocument.AddMainDocumentPart();
            part.Document = new Document(new Body());
        }

        public TestsBase(ITestOutputHelper log)
        {
            this.Log = log;
            this._validator = new OpenXmlValidator();
        }

        protected readonly ITestOutputHelper Log;
        private readonly OpenXmlValidator _validator;

        private static readonly Lazy<string> s_tempDir = new(() =>
        {
            var dir = new DirectoryInfo("./../../../../temp");
            if (dir.Exists)
                dir.Delete(true);
            dir.Create();
            return dir.FullName;
        });

        protected static string TempDir => s_tempDir.Value;

        protected void Validate(OpenXmlPackage package, List<string> expectedErrors)
        {
            var errors = _validator.Validate(package).Where(ve =>
            {
                var found = expectedErrors.Any(xe => ve.Description.Contains(xe));
                return !found;
            }).ToList();
            
            foreach (var item in errors)
            {
                Log.WriteLine(item.Description);
            }

            Assert.Empty(errors);
        }

        protected void ValidateUniqueDocPrIds(FileInfo fi)
        {
            using var doc = WordprocessingDocument.Open(fi.FullName, false);
            var docPrIds = new HashSet<string>();
            foreach (var item in doc.MainDocumentPart.GetXDocument().Descendants(WP.docPr))
                Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
            
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
            
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
            
            if (doc.MainDocumentPart.FootnotesPart is not null)
                foreach (var item in doc.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
            
            if (doc.MainDocumentPart.EndnotesPart is not null)
                foreach (var item in doc.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                    Assert.True(docPrIds.Add(item.Attribute(NoNamespace.id).Value));
        }
    }
}
