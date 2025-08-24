// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

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

        private readonly OpenXmlValidator _validator = new();

        private static readonly Lazy<string> s_tempDir = new(() =>
        {
            var dir = new DirectoryInfo("./../../../../temp");
            if (dir.Exists)
                dir.Delete(true);
            dir.Create();
            return dir.FullName;
        });

        protected static string TempDir => s_tempDir.Value;

        protected async Task Validate(OpenXmlPackage package, List<string> expectedErrors)
        {
            var errors = _validator
                .Validate(package)
                .Where(ve =>
                {
                    var found = expectedErrors.Any(xe => ve.Description.Contains(xe));
                    return !found;
                })
                .ToList();

            foreach (var item in errors)
            {
                Console.WriteLine(item.Description);
            }

            await Assert.That(errors).IsEmpty();
        }

        protected async Task ValidateUniqueDocPrIds(FileInfo fi)
        {
            using var doc = WordprocessingDocument.Open(fi.FullName, false);
            var docPrIds = new HashSet<string>();
            foreach (var item in doc.MainDocumentPart.GetXDocument().Descendants(WP.docPr))
                await Assert.That(docPrIds.Add(item.Attribute(NoNamespace.id).Value)).IsTrue();

            foreach (var header in doc.MainDocumentPart.HeaderParts)
            foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                await Assert.That(docPrIds.Add(item.Attribute(NoNamespace.id).Value)).IsTrue();

            foreach (var footer in doc.MainDocumentPart.FooterParts)
            foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                await Assert.That(docPrIds.Add(item.Attribute(NoNamespace.id).Value)).IsTrue();

            if (doc.MainDocumentPart.FootnotesPart is not null)
                foreach (var item in doc.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                    await Assert.That(docPrIds.Add(item.Attribute(NoNamespace.id).Value)).IsTrue();

            if (doc.MainDocumentPart.EndnotesPart is not null)
                foreach (var item in doc.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                    await Assert.That(docPrIds.Add(item.Attribute(NoNamespace.id).Value)).IsTrue();
        }
    }
}
