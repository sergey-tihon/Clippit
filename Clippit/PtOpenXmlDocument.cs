// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

/*
Here is modification of a WmlDocument:
    public static WmlDocument SimplifyMarkup(WmlDocument doc, SimplifyMarkupSettings settings)
    {
        using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
        {
            using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
            {
                SimplifyMarkup(document, settings);
            }
            return streamDoc.GetModifiedWmlDocument();
        }
    }

Here is read-only of a WmlDocument:

    public static string GetBackgroundColor(WmlDocument doc)
    {
        using (OpenXmlMemoryStreamDocument streamDoc = new OpenXmlMemoryStreamDocument(doc))
        using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
        {
            XDocument mainDocument = document.MainDocumentPart.GetXDocument();
            XElement backgroundElement = mainDocument.Descendants(W.background).FirstOrDefault();
            return (backgroundElement == null) ? string.Empty : backgroundElement.Attribute(W.color).Value;
        }
    }

Here is creating a new WmlDocument:

    private OpenXmlPowerToolsDocument CreateSplitDocument(WordprocessingDocument source, List<XElement> contents, string newFileName)
    {
        using (OpenXmlMemoryStreamDocument streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument())
        {
            using (WordprocessingDocument document = streamDoc.GetWordprocessingDocument())
            {
                DocumentBuilder.FixRanges(source.MainDocumentPart.GetXDocument(), contents);
                PowerToolsExtensions.SetContent(document, contents);
            }
            OpenXmlPowerToolsDocument newDoc = streamDoc.GetModifiedDocument();
            newDoc.FileName = newFileName;
            return newDoc;
        }
    }
*/

using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;
using Clippit.Excel;
using Clippit.PowerPoint;
using Clippit.Word;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit
{
    public class PowerToolsDocumentException : Exception
    {
        public PowerToolsDocumentException(string message)
            : base(message) { }
    }

    public class PowerToolsInvalidDataException : Exception
    {
        public PowerToolsInvalidDataException(string message)
            : base(message) { }
    }

    public class OpenXmlPowerToolsDocument
    {
        public string FileName { get; set; }
        public byte[] DocumentByteArray { get; set; }

        public static OpenXmlPowerToolsDocument FromFileName(string fileName)
        {
            var bytes = File.ReadAllBytes(fileName);
            Type type;
            try
            {
                type = GetDocumentType(bytes);
            }
            catch (FileFormatException)
            {
                throw new PowerToolsDocumentException("Not an Open XML document.");
            }
            if (type == typeof(WordprocessingDocument))
                return new WmlDocument(fileName, bytes);
            if (type == typeof(SpreadsheetDocument))
                return new SmlDocument(fileName, bytes);
            if (type == typeof(PresentationDocument))
                return new PmlDocument(fileName, bytes);
            if (type == typeof(Package))
            {
                return new OpenXmlPowerToolsDocument(bytes) { FileName = fileName };
            }
            throw new PowerToolsDocumentException("Not an Open XML document.");
        }

        public static OpenXmlPowerToolsDocument FromDocument(OpenXmlPowerToolsDocument doc)
        {
            var type = doc.GetDocumentType();
            if (type == typeof(WordprocessingDocument))
                return new WmlDocument(doc);
            if (type == typeof(SpreadsheetDocument))
                return new SmlDocument(doc);
            if (type == typeof(PresentationDocument))
                return new PmlDocument(doc);
            return null; // This should not be possible from a valid OpenXmlPowerToolsDocument object
        }

        public OpenXmlPowerToolsDocument(OpenXmlPowerToolsDocument original)
        {
            DocumentByteArray = new byte[original.DocumentByteArray.Length];
            Array.Copy(original.DocumentByteArray, DocumentByteArray, original.DocumentByteArray.Length);
            FileName = original.FileName;
        }

        public OpenXmlPowerToolsDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
        {
            if (convertToTransitional)
            {
                ConvertToTransitional(original.FileName, original.DocumentByteArray);
            }
            else
            {
                DocumentByteArray = new byte[original.DocumentByteArray.Length];
                Array.Copy(original.DocumentByteArray, DocumentByteArray, original.DocumentByteArray.Length);
                FileName = original.FileName;
            }
        }

        public OpenXmlPowerToolsDocument(string fileName)
        {
            this.FileName = fileName;
            DocumentByteArray = File.ReadAllBytes(fileName);
        }

        public OpenXmlPowerToolsDocument(string fileName, bool convertToTransitional)
        {
            this.FileName = fileName;

            if (convertToTransitional)
            {
                var tempByteArray = File.ReadAllBytes(fileName);
                ConvertToTransitional(fileName, tempByteArray);
            }
            else
            {
                this.FileName = fileName;
                DocumentByteArray = File.ReadAllBytes(fileName);
            }
        }

        private void ConvertToTransitional(string fileName, byte[] tempByteArray)
        {
            Type type;
            try
            {
                type = GetDocumentType(tempByteArray);
            }
            catch (FileFormatException)
            {
                throw new PowerToolsDocumentException("Not an Open XML document.");
            }

            using var ms = new MemoryStream();
            ms.Write(tempByteArray, 0, tempByteArray.Length);
            if (type == typeof(WordprocessingDocument))
            {
                using var sDoc = WordprocessingDocument.Open(ms, true);
                // following code forces the SDK to serialize
                foreach (var part in sDoc.Parts)
                {
                    try
                    {
                        var _ = part.OpenXmlPart.RootElement;
                    }
                    catch (Exception) { }
                }
            }
            else if (type == typeof(SpreadsheetDocument))
            {
                using var sDoc = SpreadsheetDocument.Open(ms, true);
                // following code forces the SDK to serialize
                foreach (var part in sDoc.Parts)
                {
                    try
                    {
                        var z = part.OpenXmlPart.RootElement;
                    }
                    catch (Exception) { }
                }
            }
            else if (type == typeof(PresentationDocument))
            {
                using var sDoc = PresentationDocument.Open(ms, true);
                // following code forces the SDK to serialize
                foreach (var part in sDoc.Parts)
                {
                    try
                    {
                        var z = part.OpenXmlPart.RootElement;
                    }
                    catch (Exception) { }
                }
            }
            this.FileName = fileName;
            DocumentByteArray = ms.ToArray();
        }

        public OpenXmlPowerToolsDocument(byte[] byteArray)
        {
            DocumentByteArray = new byte[byteArray.Length];
            Array.Copy(byteArray, DocumentByteArray, byteArray.Length);
            this.FileName = null;
        }

        public OpenXmlPowerToolsDocument(byte[] byteArray, bool convertToTransitional)
        {
            if (convertToTransitional)
            {
                ConvertToTransitional(null, byteArray);
            }
            else
            {
                DocumentByteArray = new byte[byteArray.Length];
                Array.Copy(byteArray, DocumentByteArray, byteArray.Length);
                this.FileName = null;
            }
        }

        public OpenXmlPowerToolsDocument(string fileName, MemoryStream memStream)
        {
            FileName = fileName;
            DocumentByteArray = memStream.ToArray();
        }

        public OpenXmlPowerToolsDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
        {
            if (convertToTransitional)
            {
                ConvertToTransitional(fileName, memStream.ToArray());
            }
            else
            {
                FileName = fileName;
                DocumentByteArray = memStream.ToArray();
            }
        }

        public string GetName()
        {
            if (FileName is null)
                return "Unnamed Document";
            var file = new FileInfo(FileName);
            return file.Name;
        }

        public void SaveAs(string fileName)
        {
            File.WriteAllBytes(fileName, DocumentByteArray);
        }

        public void Save()
        {
            if (FileName is null)
                throw new InvalidOperationException(
                    "Attempting to Save a document that has no file name.  Use SaveAs instead."
                );
            File.WriteAllBytes(FileName, DocumentByteArray);
        }

        public void WriteByteArray(Stream stream)
        {
            stream.Write(DocumentByteArray, 0, DocumentByteArray.Length);
        }

        public Type GetDocumentType()
        {
            return GetDocumentType(DocumentByteArray);
        }

        private static Type GetDocumentType(byte[] bytes)
        {
            using var stream = new MemoryStream();
            stream.Write(bytes, 0, bytes.Length);
            using var package = Package.Open(stream, FileMode.Open);
            var relationship =
                package
                    .GetRelationshipsByType(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    )
                    .FirstOrDefault()
                ?? package
                    .GetRelationshipsByType("http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
                    .FirstOrDefault();
            if (relationship is null)
                return null;

            var part = package.GetPart(PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri));
            switch (part.ContentType)
            {
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml":
                case "application/vnd.ms-word.document.macroEnabled.main+xml":
                case "application/vnd.ms-word.template.macroEnabledTemplate.main+xml":
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml":
                    return typeof(WordprocessingDocument);
                case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":
                case "application/vnd.ms-excel.sheet.macroEnabled.main+xml":
                case "application/vnd.ms-excel.template.macroEnabled.main+xml":
                case "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml":
                    return typeof(SpreadsheetDocument);
                case "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml":
                case "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml":
                case "application/vnd.ms-powerpoint.template.macroEnabled.main+xml":
                case "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml":
                case "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml":
                case "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml":
                    return typeof(PresentationDocument);
            }
            return typeof(Package);
        }

        public static void SavePartAs(OpenXmlPart part, string filePath)
        {
            //Stream partStream = part.GetStream(FileMode.Open, FileAccess.Read);
            //byte[] partContent = new byte[partStream.Length];
            //partStream.Read(partContent, 0, (int)partStream.Length);

            //File.WriteAllBytes(filePath, partContent);

            using var fileStream = File.Create(filePath);
            using var partStream = part.GetStream(FileMode.Open, FileAccess.Read);
            partStream.CopyTo(fileStream);
        }
    }

    public class OpenXmlMemoryStreamDocument : IDisposable
    {
        private readonly OpenXmlPowerToolsDocument _document;
        private MemoryStream _docMemoryStream;
        private Package _docPackage;

#pragma warning disable IDISP003
        public OpenXmlMemoryStreamDocument(OpenXmlPowerToolsDocument doc)
        {
            _document = doc;
            _docMemoryStream = new MemoryStream();
            _docMemoryStream.Write(doc.DocumentByteArray, 0, doc.DocumentByteArray.Length);
            try
            {
                _docPackage = Package.Open(_docMemoryStream, FileMode.Open);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }
#pragma warning restore IDISP003

        private OpenXmlMemoryStreamDocument(MemoryStream stream)
        {
            _docMemoryStream = stream;
            try
            {
                _docPackage = Package.Open(_docMemoryStream, FileMode.Open);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public static OpenXmlMemoryStreamDocument CreateWordprocessingDocument(MemoryStream stream)
        {
            stream ??= new MemoryStream();
            using (
                var doc = WordprocessingDocument.Create(
                    stream,
                    DocumentFormat.OpenXml.WordprocessingDocumentType.Document
                )
            )
            {
                doc.AddMainDocumentPart();
                doc.MainDocumentPart.PutXDocument(
                    new XDocument(
                        new XElement(
                            W.document,
                            new XAttribute(XNamespace.Xmlns + "w", W.w),
                            new XAttribute(XNamespace.Xmlns + "r", R.r),
                            new XElement(W.body)
                        )
                    )
                );
            }

            return new OpenXmlMemoryStreamDocument(stream);
        }

        public static OpenXmlMemoryStreamDocument CreateSpreadsheetDocument(MemoryStream stream)
        {
            stream ??= new MemoryStream();
            using (
                var doc = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook)
            )
            {
                doc.AddWorkbookPart();
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                XNamespace relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                doc.WorkbookPart.PutXDocument(
                    new XDocument(
                        new XElement(
                            ns + "workbook",
                            new XAttribute("xmlns", ns),
                            new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                            new XElement(ns + "sheets")
                        )
                    )
                );
            }

            return new OpenXmlMemoryStreamDocument(stream);
        }

        public static OpenXmlMemoryStreamDocument CreatePresentationDocument(MemoryStream stream = null)
        {
            stream ??= new MemoryStream();
            using (
                var doc = PresentationDocument.Create(
                    stream,
                    DocumentFormat.OpenXml.PresentationDocumentType.Presentation
                )
            )
            {
                doc.AddPresentationPart();
                XNamespace ns = "http://schemas.openxmlformats.org/presentationml/2006/main";
                XNamespace relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                XNamespace drawingns = "http://schemas.openxmlformats.org/drawingml/2006/main";
                doc.PresentationPart.PutXDocument(
                    new XDocument(
                        new XElement(
                            ns + "presentation",
                            new XAttribute(XNamespace.Xmlns + "a", drawingns),
                            new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                            new XAttribute(XNamespace.Xmlns + "p", ns),
                            new XElement(ns + "sldMasterIdLst"),
                            new XElement(ns + "sldIdLst"),
                            new XElement(
                                ns + "notesSz",
                                new XAttribute("cx", "6858000"),
                                new XAttribute("cy", "9144000")
                            )
                        )
                    )
                );
            }

            return new OpenXmlMemoryStreamDocument(stream);
        }

        public static OpenXmlMemoryStreamDocument CreatePackage()
        {
            var stream = new MemoryStream();
            using (var package = Package.Open(stream, FileMode.Create))
                package.Close();
            return new OpenXmlMemoryStreamDocument(stream);
        }

        public Package GetPackage()
        {
            return _docPackage;
        }

        public WordprocessingDocument GetWordprocessingDocument(OpenSettings openSettings = null)
        {
            try
            {
                if (GetDocumentType() != typeof(WordprocessingDocument))
                    throw new PowerToolsDocumentException("Not a Wordprocessing document.");
                return openSettings is null
                    ? WordprocessingDocument.Open(_docPackage)
                    : WordprocessingDocument.Open(_docPackage, openSettings);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public SpreadsheetDocument GetSpreadsheetDocument(OpenSettings openSettings = null)
        {
            try
            {
                if (GetDocumentType() != typeof(SpreadsheetDocument))
                    throw new PowerToolsDocumentException("Not a Spreadsheet document.");
                return openSettings is null
                    ? SpreadsheetDocument.Open(_docPackage)
                    : SpreadsheetDocument.Open(_docPackage, openSettings);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public PresentationDocument GetPresentationDocument(OpenSettings openSettings = null)
        {
            try
            {
                if (GetDocumentType() != typeof(PresentationDocument))
                    throw new PowerToolsDocumentException("Not a Presentation document.");
                return openSettings is null
                    ? PresentationDocument.Open(_docPackage)
                    : PresentationDocument.Open(_docPackage, openSettings);
            }
            catch (Exception e)
            {
                throw new PowerToolsDocumentException(e.Message);
            }
        }

        public Type GetDocumentType()
        {
            var relationship =
                _docPackage
                    .GetRelationshipsByType(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    )
                    .FirstOrDefault()
                ?? _docPackage
                    .GetRelationshipsByType("http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
                    .FirstOrDefault();
            if (relationship == null)
                throw new PowerToolsDocumentException("Not an Open XML Document.");
            var part = _docPackage.GetPart(
                PackUriHelper.ResolvePartUri(relationship.SourceUri, relationship.TargetUri)
            );
            switch (part.ContentType)
            {
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml":
                case "application/vnd.ms-word.document.macroEnabled.main+xml":
                case "application/vnd.ms-word.template.macroEnabledTemplate.main+xml":
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml":
                    return typeof(WordprocessingDocument);
                case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml":
                case "application/vnd.ms-excel.sheet.macroEnabled.main+xml":
                case "application/vnd.ms-excel.template.macroEnabled.main+xml":
                case "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml":
                    return typeof(SpreadsheetDocument);
                case "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml":
                case "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml":
                case "application/vnd.ms-powerpoint.template.macroEnabled.main+xml":
                case "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml":
                case "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml":
                case "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml":
                    return typeof(PresentationDocument);
            }
            return null;
        }

#pragma warning disable IDISP003
        public void ClosePackage()
        {
            _docPackage.Close();
            _docPackage = null;
        }

        public OpenXmlPowerToolsDocument GetModifiedDocument()
        {
            ClosePackage();
            return new OpenXmlPowerToolsDocument(_document?.FileName, _docMemoryStream);
        }

        public WmlDocument GetModifiedWmlDocument()
        {
            ClosePackage();
            return new WmlDocument(_document?.FileName, _docMemoryStream);
        }

        public SmlDocument GetModifiedSmlDocument()
        {
            ClosePackage();
            return new SmlDocument(_document?.FileName, _docMemoryStream);
        }

        public PmlDocument GetModifiedPmlDocument()
        {
            ClosePackage();
            return new PmlDocument(_document?.FileName, _docMemoryStream);
        }

        private bool _disposedValue;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~OpenXmlMemoryStreamDocument()
        {
            Dispose(false);
        }

        private void Dispose(bool disposing)
        {
            if (_disposedValue)
                return;

            if (disposing)
            {
                _docPackage?.Close();
                _docMemoryStream?.Dispose();
                _docPackage = null;
                _docMemoryStream = null;
            }
            _disposedValue = true;
        }
#pragma warning restore IDISP003
    }
}
