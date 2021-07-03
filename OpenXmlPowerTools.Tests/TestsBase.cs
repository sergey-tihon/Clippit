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
        private const WordprocessingDocumentType DocumentType = WordprocessingDocumentType.Document;

        protected static void CreateEmptyWordprocessingDocument(Stream stream)
        {
            using var wordDocument = WordprocessingDocument.Create(stream, DocumentType);
            var part = wordDocument.AddMainDocumentPart();
            part.Document = new Document(new Body());
        }

        protected TestsBase(ITestOutputHelper log)
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

        public static string TempDir => s_tempDir.Value;
                
        protected void Validate(SpreadsheetDocument sDoc)
        {
            var errors = _validator.Validate(sDoc)
                .Where(ve => !s_spreadsheetExpectedErrors.Contains(ve.Description))
                .ToList();

            // if a test fails validation post-processing, then can use this code to determine the SDK
            // validation error(s).
            foreach (var item in errors)
            {
                Log.WriteLine(item.Description);
            }
            
            Assert.Empty(errors);
        }

        private static readonly List<string> s_spreadsheetExpectedErrors = new()
        {
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        };
    }
}
