// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

#define TestForUnsupportedDocuments
#define MergeStylesWithSameNames

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Clippit.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Word
{
    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        public IEnumerable<WmlDocument> SplitOnSections()
        {
            return Word.DocumentBuilder.SplitOnSections(this);
        }
    }

    public interface ISource : ICloneable
    {
        WmlDocument WmlDocument { get; set; }

        bool KeepSections { get; set; }
        public bool DiscardHeadersAndFootersInKeptSections { get; set; }

        string InsertId { get; set; }

        IEnumerable<XElement> GetElements(WordprocessingDocument document);
    }

    [Serializable]
    public class Source : ISource
    {
        public WmlDocument WmlDocument
        {
            get => _wmlDocument;
            set => _wmlDocument = value;
        }

        [NonSerialized] private WmlDocument _wmlDocument;


        public int Start { get; set; }
        public int Count { get; set; }
        public bool KeepSections { get; set; }
        public bool DiscardHeadersAndFootersInKeptSections { get; set; }
        public string InsertId { get; set; }

        public Source(string fileName)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = null;
        }

        public Source(WmlDocument source)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = null;
        }

        public Source(string fileName, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, bool keepSections)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, string insertId)
        {
            WmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(string fileName, int start, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, int start, bool keepSections)
        {
            WmlDocument = source;
            Start = start;
            Count = int.MaxValue;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, int start, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, int start, string insertId)
        {
            WmlDocument = source;
            Start = start;
            Count = int.MaxValue;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(string fileName, int start, int count, bool keepSections)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = count;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(WmlDocument source, int start, int count, bool keepSections)
        {
            WmlDocument = source;
            Start = start;
            Count = count;
            KeepSections = keepSections;
            InsertId = null;
        }

        public Source(string fileName, int start, int count, string insertId)
        {
            WmlDocument = new WmlDocument(fileName);
            Start = start;
            Count = count;
            KeepSections = false;
            InsertId = insertId;
        }

        public Source(WmlDocument source, int start, int count, string insertId)
        {
            WmlDocument = source;
            Start = start;
            Count = count;
            KeepSections = false;
            InsertId = insertId;
        }

        public IEnumerable<XElement> GetElements(WordprocessingDocument document)
        {
            var body = document.MainDocumentPart.GetXDocument().Root?.Element(W.body);

            if (body is null)
                throw new DocumentBuilderException(
                    "Unsupported document - contains no body element in the correct namespace");

            return body
                .Elements()
                .Skip(Start)
                .Take(Count)
                .ToList();
        }

        public object Clone() =>
            new Source(WmlDocument, Start, Count, KeepSections)
            {
                DiscardHeadersAndFootersInKeptSections = DiscardHeadersAndFootersInKeptSections,
                InsertId = InsertId
            };
    }

    [Serializable]
    public class TableCellSource : ISource
    {
        public TableCellSource()
        {
            KeepSections = false;
            DiscardHeadersAndFootersInKeptSections = false;
        }
        
        public TableCellSource(WmlDocument source) : this()
        {
            WmlDocument = source;
        }

        public WmlDocument WmlDocument
        {
            get => _wmlDocument;
            set => _wmlDocument = value;
        }

        [NonSerialized] private WmlDocument _wmlDocument;


        public bool KeepSections { get; set; }
        public bool DiscardHeadersAndFootersInKeptSections { get; set; }

        public string InsertId { get; set; }


        public int TableElementIndex { get; set; }

        public int RowIndex { get; set; }

        public int CellIndex { get; set; }

        public int CellContentStart { get; set; }

        public int CellContentCount { get; set; }


        public IEnumerable<XElement> GetElements(WordprocessingDocument document)
        {
            var body = document.MainDocumentPart.GetXDocument().Root?.Element(W.body);

            if (body is null)
            {
                throw new DocumentBuilderException(
                    "Unsupported document - contains no body element in the correct namespace");
            }

            var table = body.Elements().Skip(TableElementIndex).FirstOrDefault();
            if (table is null || table.Name != W.tbl)
            {
                throw new DocumentBuilderException(
                    $"Invalid {nameof(TableCellSource)} - element {TableElementIndex} is '{table?.Name}' but expected {W.tbl}");
            }

            var row = table.Elements(W.tr).Skip(RowIndex).FirstOrDefault();
            if (row is null)
            {
                throw new DocumentBuilderException(
                    $"Invalid {nameof(TableCellSource)} - row {RowIndex} does not exist");

            }

            var cell = row.Elements(W.tc).Skip(CellIndex).FirstOrDefault();
            if (cell is null)
            {
                throw new DocumentBuilderException(
                    $"Invalid {nameof(TableCellSource)} - cell {CellIndex} in the row {RowIndex} does not exist");

            }

            return cell
                .Elements()
                .Skip(CellContentStart)
                .Take(CellContentCount)
                .ToList();
        }

        public object Clone() =>
            new TableCellSource
            {
                WmlDocument = WmlDocument,
                KeepSections = KeepSections,
                DiscardHeadersAndFootersInKeptSections = DiscardHeadersAndFootersInKeptSections,
                InsertId = InsertId,
                TableElementIndex = TableElementIndex,
                RowIndex = RowIndex,
                CellIndex = CellIndex,
                CellContentStart = CellContentStart,
                CellContentCount = CellContentCount
            };
    }

    public class DocumentBuilderSettings
    {
        public HashSet<string> CustomXmlGuidList = null;
        public bool NormalizeStyleIds = false;
    }

    public static class DocumentBuilder
    {
        public static void BuildDocument(List<ISource> sources, string fileName)
        {
            using var streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
            using (var output = streamDoc.GetWordprocessingDocument())
            {
                BuildDocument(sources, output, new DocumentBuilderSettings());
                output.Close();
            }
            streamDoc.GetModifiedDocument().SaveAs(fileName);
        }

        public static void BuildDocument(List<ISource> sources, string fileName, DocumentBuilderSettings settings)
        {
            using var streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
            using (var output = streamDoc.GetWordprocessingDocument())
            {
                BuildDocument(sources, output, settings);
                output.Close();
            }
            streamDoc.GetModifiedDocument().SaveAs(fileName);
        }

        public static WmlDocument BuildDocument(List<ISource> sources)
        {
            using var streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
            using (var output = streamDoc.GetWordprocessingDocument())
            {
                BuildDocument(sources, output, new DocumentBuilderSettings());
                output.Close();
            }
            return streamDoc.GetModifiedWmlDocument();
        }

        public static WmlDocument BuildDocument(List<ISource> sources, DocumentBuilderSettings settings)
        {
            using var streamDoc = OpenXmlMemoryStreamDocument.CreateWordprocessingDocument();
            using (var output = streamDoc.GetWordprocessingDocument())
            {
                BuildDocument(sources, output, settings);
                output.Close();
            }
            return streamDoc.GetModifiedWmlDocument();
        }

        private struct TempSource
        {
            public int Start;
            public int Count;
        };

        private class Atbi
        {
            public XElement BlockLevelContent;
            public int Index;
        }

        private class Atbid
        {
            public XElement BlockLevelContent;
            public int Index;
            public int Div;
        }

        private const string Yes = "yes";
        private const string Utf8 = "UTF-8";
        private const string OnePointZero = "1.0";

        public static IEnumerable<WmlDocument> SplitOnSections(WmlDocument doc)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(doc);
            using var document = streamDoc.GetWordprocessingDocument();
            var mainDocument = document.MainDocumentPart.GetXDocument();
            var divs = mainDocument
                .Root
                .Element(W.body)
                .Elements()
                .Select((p, i) => new Atbi
                {
                    BlockLevelContent = p,
                    Index = i,
                })
                .Rollup(new Atbid
                    {
                        BlockLevelContent = (XElement)null,
                        Index = -1,
                        Div = 0,
                    },
                    (b, p) =>
                    {
                        var elementBefore = b.BlockLevelContent
                            .SiblingsBeforeSelfReverseDocumentOrder()
                            .FirstOrDefault();
                        if (elementBefore != null && elementBefore.Descendants(W.sectPr).Any())
                            return new Atbid
                            {
                                BlockLevelContent = b.BlockLevelContent,
                                Index = b.Index,
                                Div = p.Div + 1,
                            };
                        return new Atbid
                        {
                            BlockLevelContent = b.BlockLevelContent,
                            Index = b.Index,
                            Div = p.Div,
                        };
                    });
            var groups = divs
                .GroupAdjacent(b => b.Div);
            var tempSourceList = groups
                .Select(g => new TempSource
                {
                    Start = g.First().Index,
                    Count = g.Count(),
                })
                .ToList();
            foreach (var ts in tempSourceList)
            {
                var sources = new List<ISource>
                {
                    new Source(doc, ts.Start, ts.Count, true)
                };
                var newDoc = BuildDocument(sources);
                newDoc = AdjustSectionBreak(newDoc);
                yield return newDoc;
            }
        }

        private static WmlDocument AdjustSectionBreak(WmlDocument doc)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(doc);
            using (var document = streamDoc.GetWordprocessingDocument())
            {
                var mainXDoc = document.MainDocumentPart.GetXDocument();
                var lastElement = mainXDoc.Root
                    .Element(W.body)
                    .Elements()
                    .LastOrDefault();
                if (lastElement != null)
                {
                    if (lastElement.Name != W.sectPr &&
                        lastElement.Descendants(W.sectPr).Any())
                    {
                        mainXDoc.Root.Element(W.body).Add(lastElement.Descendants(W.sectPr).First());
                        lastElement.Descendants(W.sectPr).Remove();
                        if (!lastElement.Elements()
                            .Where(e => e.Name != W.pPr)
                            .Any())
                            lastElement.Remove();
                        document.MainDocumentPart.PutXDocument();
                    }
                }
            }
            return streamDoc.GetModifiedWmlDocument();
        }

        private static void BuildDocument(List<ISource> sources, WordprocessingDocument output, DocumentBuilderSettings settings)
        {
            var wmlGlossaryDocument = CoalesceGlossaryDocumentParts(sources, settings);

            if (RelationshipMarkup == null)
                InitRelationshipMarkup();

            // This list is used to eliminate duplicate images
            var images = new List<ImageData>();
            var mainPart = output.MainDocumentPart.GetXDocument();
            mainPart.Declaration.Standalone = Yes;
            mainPart.Declaration.Encoding = Utf8;
            mainPart.Root.ReplaceWith(
                new XElement(W.document, NamespaceAttributes,
                    new XElement(W.body)));
            if (sources.Count > 0)
            {
                // the following function makes sure that for a given style name, the same style ID is used for all documents.
                if (settings != null && settings.NormalizeStyleIds)
                    sources = NormalizeStyleNamesAndIds(sources);

                using (var streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument))
                using (var doc = streamDoc.GetWordprocessingDocument())
                {
                    CopyStartingParts(doc, output, images);
                    CopySpecifiedCustomXmlParts(doc, output, settings);
                }

                var sourceNum2 = 0;
                foreach (var source in sources)
                {
                    if (source.InsertId != null)
                    {
                        while (true)
                        {
#if false
                            modify AppendDocument so that it can take a part.
                            for each in main document part, header parts, footer parts
                                are there any PtOpenXml.Insert elements in any of them?
                            if so, then open and process all.
#endif
                            var foundInMainDocPart = false;
                            var mainXDoc = output.MainDocumentPart.GetXDocument();
                            if (mainXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId))
                                foundInMainDocPart = true;
                            if (foundInMainDocPart)
                            {
                                using var streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument);
                                using var doc = streamDoc.GetWordprocessingDocument();
#if TestForUnsupportedDocuments
                                // throws exceptions if a document contains unsupported content
                                TestForUnsupportedDocument(doc, sources.IndexOf(source));
#endif
                                if (foundInMainDocPart)
                                {
                                    if (source.KeepSections && source.DiscardHeadersAndFootersInKeptSections)
                                        RemoveHeadersAndFootersFromSections(doc);
                                    else if (source.KeepSections)
                                        ProcessSectionsForLinkToPreviousHeadersAndFooters(doc);

                                    var contents = source.GetElements(doc).ToList();

                                    try
                                    {
                                        AppendDocument(doc, output, contents, source.KeepSections, source.InsertId, images);
                                    }
                                    catch (DocumentBuilderInternalException dbie)
                                    {
                                        if (dbie.Message.Contains("{0}"))
                                            throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum2));
                                        throw;
                                    }
                                }
                            }
                            else
                                break;
                        }
                    }
                    else
                    {
                        using var streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument);
                        using var doc = streamDoc.GetWordprocessingDocument();
#if TestForUnsupportedDocuments
                        // throws exceptions if a document contains unsupported content
                        TestForUnsupportedDocument(doc, sources.IndexOf(source));
#endif
                        if (source.KeepSections && source.DiscardHeadersAndFootersInKeptSections)
                            RemoveHeadersAndFootersFromSections(doc);
                        else if (source.KeepSections)
                            ProcessSectionsForLinkToPreviousHeadersAndFooters(doc);

                        var contents = source.GetElements(doc).ToList();
                        try
                        {
                            AppendDocument(doc, output, contents, source.KeepSections, null, images);
                        }
                        catch (DocumentBuilderInternalException dbie)
                        {
                            if (dbie.Message.Contains("{0}"))
                                throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum2));
                            throw;
                        }
                    }
                    ++sourceNum2;
                }
                if (!sources.Any(s => s.KeepSections))
                {
                    using var streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument);
                    using var doc = streamDoc.GetWordprocessingDocument();
                    var body = doc.MainDocumentPart.GetXDocument().Root.Element(W.body);

                    if (body != null && body.Elements().Any())
                    {
                        var sectPr = body.Elements().LastOrDefault();
                        if (sectPr != null && sectPr.Name == W.sectPr)
                        {
                            AddSectionAndDependencies(doc, output, sectPr, images);
                            body.Add(sectPr);
                        }
                    }
                }
                else
                {
                    FixUpSectionProperties(output);

                    // Any sectPr elements that do not have headers and footers should take their headers and footers from the *next* section,
                    // i.e. from the running section.
                    var mxd = output.MainDocumentPart.GetXDocument();
                    var sections = mxd.Descendants(W.sectPr).Reverse().ToList();

                    var cachedHeaderFooter = new[]
                    {
                        new CachedHeaderFooter { Ref = W.headerReference, Type = "first" },
                        new CachedHeaderFooter { Ref = W.headerReference, Type = "even" },
                        new CachedHeaderFooter { Ref = W.headerReference, Type = "default" },
                        new CachedHeaderFooter { Ref = W.footerReference, Type = "first" },
                        new CachedHeaderFooter { Ref = W.footerReference, Type = "even" },
                        new CachedHeaderFooter { Ref = W.footerReference, Type = "default" },
                    };

                    var firstSection = true;
                    foreach (var sect in sections)
                    {
                        if (firstSection)
                        {
                            foreach (var hf in cachedHeaderFooter)
                            {
                                var referenceElement = sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type);
                                if (referenceElement != null)
                                    hf.CachedPartRid = (string)referenceElement.Attribute(R.id);
                            }
                            firstSection = false;
                            continue;
                        }

                        CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.headerReference, "first");
                        CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.headerReference, "even");
                        CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.headerReference, "default");
                        CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.footerReference, "first");
                        CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.footerReference, "even");
                        CopyOrCacheHeaderOrFooter(output, cachedHeaderFooter, sect, W.footerReference, "default");
                    }
                }

                // Now can process PtOpenXml:Insert elements in headers / footers
                var sourceNum = 0;
                foreach (var source in sources)
                {
                    if (source.InsertId != null)
                    {
                        while (true)
                        {
#if false
                            this uses an overload of AppendDocument that takes a part.
                            for each in main document part, header parts, footer parts
                                are there any PtOpenXml.Insert elements in any of them?
                            if so, then open and process all.
#endif
                            var foundInHeadersFooters = false;
                            if (output.MainDocumentPart.HeaderParts.Any(hp =>
                            {
                                var hpXDoc = hp.GetXDocument();
                                return hpXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId);
                            }))
                                foundInHeadersFooters = true;
                            if (output.MainDocumentPart.FooterParts.Any(fp =>
                            {
                                var hpXDoc = fp.GetXDocument();
                                return hpXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId);
                            }))
                                foundInHeadersFooters = true;

                            if (foundInHeadersFooters)
                            {
                                using var streamDoc = new OpenXmlMemoryStreamDocument(source.WmlDocument);
                                using var doc = streamDoc.GetWordprocessingDocument();
#if TestForUnsupportedDocuments
                                // throws exceptions if a document contains unsupported content
                                TestForUnsupportedDocument(doc, sources.IndexOf(source));
#endif
                                var partList = output.MainDocumentPart.HeaderParts.Concat(output.MainDocumentPart.FooterParts.Cast<OpenXmlPart>()).ToList();
                                foreach (var part in partList)
                                {
                                    var partXDoc = part.GetXDocument();
                                    if (!partXDoc.Descendants(PtOpenXml.Insert).Any(d => (string)d.Attribute(PtOpenXml.Id) == source.InsertId))
                                        continue;
                                    var contents = source.GetElements(doc).ToList();

                                    try
                                    {
                                        AppendDocument(doc, output, part, contents, source.KeepSections, source.InsertId, images);
                                    }
                                    catch (DocumentBuilderInternalException dbie)
                                    {
                                        if (dbie.Message.Contains("{0}"))
                                            throw new DocumentBuilderException(string.Format(dbie.Message, sourceNum));
                                        throw;
                                    }
                                }
                            }
                            else
                                break;
                        }
                    }
                    ++sourceNum;
                }
                if (sources.Any(s => s.KeepSections) && !output.MainDocumentPart.GetXDocument().Root.Descendants(W.sectPr).Any())
                {
                    using var streamDoc = new OpenXmlMemoryStreamDocument(sources[0].WmlDocument);
                    using var doc = streamDoc.GetWordprocessingDocument();
                    var body = doc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                    var sectPr = body.Elements().LastOrDefault();
                    if (sectPr != null && sectPr.Name == W.sectPr)
                    {
                        AddSectionAndDependencies(doc, output, sectPr, images);
                        body.Add(sectPr);
                    }
                }
                AdjustDocPrIds(output);
            }

            if (wmlGlossaryDocument != null)
                WriteGlossaryDocumentPart(wmlGlossaryDocument, output, images);

            foreach (var part in output.GetAllParts())
                if (part.Annotation<XDocument>() != null)
                    part.PutXDocument();
        }

        // there are two scenarios that need to be handled
        // - if I find a style name that maps to a style ID different from one already mapped
        // - if a style name maps to a style ID that is already used for a different style
        // - then need to correct things
        //   - make a complete list of all things that need to be changed, for every correction
        //   - do the corrections all at one
        //   - mark the document as changed, and change it in the sources.
        private static List<ISource> NormalizeStyleNamesAndIds(List<ISource> sources)
        {
            var styleNameMap = new Dictionary<string, string>();
            var styleIds = new HashSet<string>();
            var newSources = new List<ISource>();

            foreach (var src in sources)
            {
                var newSrc = AddAndRectify(src, styleNameMap, styleIds);
                newSources.Add(newSrc);
            }
            return newSources;
        }

        private static ISource AddAndRectify(ISource src, Dictionary<string, string> styleNameMap, HashSet<string> styleIds)
        {
            var modified = false;
            using var ms = new MemoryStream();
            ms.Write(src.WmlDocument.DocumentByteArray, 0, src.WmlDocument.DocumentByteArray.Length);
            using (var wDoc = WordprocessingDocument.Open(ms, true))
            {
                var correctionList = new Dictionary<string, string>();
                var thisStyleNameMap = GetStyleNameMap(wDoc);
                foreach (var pair in thisStyleNameMap)
                {
                    var styleName = pair.Key;
                    var styleId = pair.Value;
                    // if the styleNameMap does not contain an entry for this name
                    if (!styleNameMap.ContainsKey(styleName))
                    {
                        // if the id is already used
                        if (styleIds.Contains(styleId))
                        {
                            // this style uses a styleId that is used for another style.
                            // randomly generate new styleId
                            while (true)
                            {
                                var newStyleId = GenStyleIdFromStyleName(styleName);
                                if (! styleIds.Contains(newStyleId))
                                {
                                    correctionList.Add(styleId, newStyleId);
                                    styleNameMap.Add(styleName, newStyleId);
                                    styleIds.Add(newStyleId);
                                    break;
                                }
                            }
                        }
                        // otherwise we just add to the styleNameMap
                        else
                        {
                            styleNameMap.Add(styleName, styleId);
                            styleIds.Add(styleId);
                        }
                    }
                    // but if the styleNameMap does contain an entry for this name
                    else
                    {
                        // if the id is the same as the existing ID, then nothing to do
                        if (styleNameMap[styleName] == styleId)
                            continue;
                        correctionList.Add(styleId, styleNameMap[styleName]);
                    }
                }
                if (correctionList.Any())
                {
                    modified = true;
                    AdjustStyleIdsForDocument(wDoc, correctionList);
                }
            }
            if (modified)
            {
                var newSrc = (ISource) src.Clone();
                newSrc.WmlDocument = new WmlDocument(src.WmlDocument.FileName, ms.ToArray());;
                return newSrc;
            }

            return src;
        }

#if false
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/@styleId
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/basedOn/@val
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/link/@val
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/next/@val

application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/@styleId
application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/basedOn/@val
application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/link/@val
application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/next/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml                         pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml                         rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml                         tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml                         pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml                         rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml                         tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml                           pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml                           rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml                           tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml                        pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml                        rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml                        tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml                           pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml                           rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml                           tblPr/tblStyle/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/lvl/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/numStyleLink/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/styleLink/@val

application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml                         settings/clickAndTypeStyle/@val

Name, not ID
===================================
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/name/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.stylesWithEffects+xml                styles/style/name/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml                           styles/style/name/@val
application/vnd.ms-word.stylesWithEffects+xml                                                       styles/style/name/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.stylesWithEffects+xml                latentStyles/lsdException/@name
application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml                           latentStyles/lsdException/@name
application/vnd.ms-word.stylesWithEffects+xml                                                       latentStyles/lsdException/@name
application/vnd.ms-word.styles.textEffects+xml                                                      latentStyles/lsdException/@name
#endif

        private static void AdjustStyleIdsForDocument(WordprocessingDocument wDoc, Dictionary<string, string> correctionList)
        {
            // update styles part
            UpdateStyleIdsForStylePart(wDoc.MainDocumentPart.StyleDefinitionsPart, correctionList);
            if (wDoc.MainDocumentPart.StylesWithEffectsPart != null)
                UpdateStyleIdsForStylePart(wDoc.MainDocumentPart.StylesWithEffectsPart, correctionList);

            // update content parts
            UpdateStyleIdsForContentPart(wDoc.MainDocumentPart, correctionList);
            foreach (var part in wDoc.MainDocumentPart.HeaderParts)
                UpdateStyleIdsForContentPart(part, correctionList);
            foreach (var part in wDoc.MainDocumentPart.FooterParts)
                UpdateStyleIdsForContentPart(part, correctionList);
            if (wDoc.MainDocumentPart.FootnotesPart != null)
                UpdateStyleIdsForContentPart(wDoc.MainDocumentPart.FootnotesPart, correctionList);
            if (wDoc.MainDocumentPart.EndnotesPart != null)
                UpdateStyleIdsForContentPart(wDoc.MainDocumentPart.EndnotesPart, correctionList);
            if (wDoc.MainDocumentPart.WordprocessingCommentsPart != null)
                UpdateStyleIdsForContentPart(wDoc.MainDocumentPart.WordprocessingCommentsPart, correctionList);
            if (wDoc.MainDocumentPart.WordprocessingCommentsExPart != null)
                UpdateStyleIdsForContentPart(wDoc.MainDocumentPart.WordprocessingCommentsExPart, correctionList);

            // update numbering part
            UpdateStyleIdsForNumberingPart(wDoc.MainDocumentPart.NumberingDefinitionsPart, correctionList);
        }

        private static void UpdateStyleIdsForNumberingPart(OpenXmlPart part, Dictionary<string, string> correctionList)
        {
#if false
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/lvl/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/numStyleLink/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml                        abstractNum/styleLink/@val
#endif
            var numXDoc = part.GetXDocument();
            var numAttributeChangeList = correctionList
                .Select(cor =>
                    new
                    {
                        NewId = cor.Value,
                        PStyleAttributesToChange = numXDoc
                            .Descendants(W.pStyle)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        NumStyleLinkAttributesToChange = numXDoc
                            .Descendants(W.numStyleLink)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        StyleLinkAttributesToChange = numXDoc
                            .Descendants(W.styleLink)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                    }
                )
                .ToList();
            foreach (var item in numAttributeChangeList)
            {
                foreach (var att in item.PStyleAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.NumStyleLinkAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.StyleLinkAttributesToChange)
                    att.Value = item.NewId;
            }
            part.PutXDocument();
        }

        private static void UpdateStyleIdsForStylePart(OpenXmlPart part, Dictionary<string, string> correctionList)
        {
#if false
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/@styleId
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/basedOn/@val
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/link/@val
application/vnd.ms-word.styles.textEffects+xml                                                      styles/style/next/@val
#endif
            var styleXDoc = part.GetXDocument();
            var styleAttributeChangeList = correctionList
                .Select(cor =>
                    new
                    {
                        NewId = cor.Value,
                        StyleIdAttributesToChange = styleXDoc
                            .Root
                            .Elements(W.style)
                            .Attributes(W.styleId)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        BasedOnAttributesToChange = styleXDoc
                            .Root
                            .Elements(W.style)
                            .Elements(W.basedOn)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        NextAttributesToChange = styleXDoc
                            .Root
                            .Elements(W.style)
                            .Elements(W.next)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        LinkAttributesToChange = styleXDoc
                            .Root
                            .Elements(W.style)
                            .Elements(W.link)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                    }
                )
                .ToList();
            foreach (var item in styleAttributeChangeList)
            {
                foreach (var att in item.StyleIdAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.BasedOnAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.NextAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.LinkAttributesToChange)
                    att.Value = item.NewId;
            }
            part.PutXDocument();
        }

        private static void UpdateStyleIdsForContentPart(OpenXmlPart part, Dictionary<string, string> correctionList)
        {
#if false
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    pPr/pStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    rPr/rStyle/@val
application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml                    tblPr/tblStyle/@val
#endif
            var xDoc = part.GetXDocument();
            var mainAttributeChangeList = correctionList
                .Select(cor =>
                    new
                    {
                        NewId = cor.Value,
                        PStyleAttributesToChange = xDoc
                            .Descendants(W.pStyle)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        RStyleAttributesToChange = xDoc
                            .Descendants(W.rStyle)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                        TblStyleAttributesToChange = xDoc
                            .Descendants(W.tblStyle)
                            .Attributes(W.val)
                            .Where(a => a.Value == cor.Key)
                            .ToList(),
                    }
                )
                .ToList();
            foreach (var item in mainAttributeChangeList)
            {
                foreach (var att in item.PStyleAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.RStyleAttributesToChange)
                    att.Value = item.NewId;
                foreach (var att in item.TblStyleAttributesToChange)
                    att.Value = item.NewId;
            }
            part.PutXDocument();
        }

        private static string GenStyleIdFromStyleName(string styleName)
        {
            var newStyleId = styleName
                .Replace("_", "")
                .Replace("#", "")
                .Replace(".", "") + ((new Random()).Next(990) + 9).ToString();
            return newStyleId;
        }

        private static Dictionary<string, string> GetStyleNameMap(WordprocessingDocument wDoc)
        {
            var sxDoc = wDoc.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
            var thisDocumentDictionary = sxDoc
                .Root
                .Elements(W.style)
                .ToDictionary(
                    z => (string)z.Elements(W.name).Attributes(W.val).FirstOrDefault(),
                    z => (string)z.Attribute(W.styleId));
            return thisDocumentDictionary;
        }

#if false
        At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
        querying or generating that markup.  This is an example of the GlossaryDocumentPart.

<w:glossaryDocument xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14">
  <w:docParts>
    <w:docPart>
      <w:docPartPr>
        <w:name w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
        <w:category>
          <w:name w:val="General" />
          <w:gallery w:val="placeholder" />
        </w:category>
        <w:types>
          <w:type w:val="bbPlcHdr" />
        </w:types>
        <w:behaviors>
          <w:behavior w:val="content" />
        </w:behaviors>
        <w:guid w:val="{13882A71-B5B7-4421-ACBB-9B61C61B3034}" />
      </w:docPartPr>
      <w:docPartBody>
        <w:p w:rsidR="00004EEA" w:rsidRDefault="00AD57F5" w:rsidP="00AD57F5">
#endif

        private static void WriteGlossaryDocumentPart(WmlDocument wmlGlossaryDocument, WordprocessingDocument output, List<ImageData> images)
        {
            using var ms = new MemoryStream();
            ms.Write(wmlGlossaryDocument.DocumentByteArray, 0, wmlGlossaryDocument.DocumentByteArray.Length);
            using var wDoc = WordprocessingDocument.Open(ms, true);
            var fromXDoc = wDoc.MainDocumentPart.GetXDocument();

            var outputGlossaryDocumentPart = output.MainDocumentPart.AddNewPart<GlossaryDocumentPart>();
            var newXDoc = new XDocument(
                new XDeclaration(OnePointZero, Utf8, Yes),
                new XElement(W.glossaryDocument,
                    NamespaceAttributes,
                    new XElement(W.docParts,
                        fromXDoc.Descendants(W.docPart))));
            outputGlossaryDocumentPart.PutXDocument(newXDoc);

            CopyGlossaryDocumentPartsToGD(wDoc, output, fromXDoc.Root.Descendants(W.docPart), images);
            CopyRelatedPartsForContentParts(wDoc.MainDocumentPart, outputGlossaryDocumentPart, new[] { fromXDoc.Root }, images);
        }

        private static WmlDocument CoalesceGlossaryDocumentParts(IEnumerable<ISource> sources, DocumentBuilderSettings settings)
        {
            var allGlossaryDocuments = sources
                .Select(s => DocumentBuilder.ExtractGlossaryDocument(s.WmlDocument))
                .Where(s => s != null)
                .Select(s => new Source(s) as ISource)
                .ToList();

            if (!allGlossaryDocuments.Any())
                return null;

            var coalescedRaw = DocumentBuilder.BuildDocument(allGlossaryDocuments);

            // now need to do some fix up
            using var ms = new MemoryStream();
            ms.Write(coalescedRaw.DocumentByteArray, 0, coalescedRaw.DocumentByteArray.Length);
            using (var wDoc = WordprocessingDocument.Open(ms, true))
            {
                var mainXDoc = wDoc.MainDocumentPart.GetXDocument();

                var body = mainXDoc.Root.Element(W.body);
                var newBody = new XElement(W.body,
                    new XElement(W.docParts,
                        body.Elements(W.docParts).Elements(W.docPart)));

                body.ReplaceWith(newBody);

                wDoc.MainDocumentPart.PutXDocument();
            }

            var coalescedGlossaryDocument = new WmlDocument("Coalesced.docx", ms.ToArray());

            return coalescedGlossaryDocument;
        }

        private static void InitRelationshipMarkup()
        {
            RelationshipMarkup = new Dictionary<XName, XName[]>
            {
                    //{ button,           new [] { image }},
                    { A.blip,             new [] { R.embed, R.link }},
                    { A.hlinkClick,       new [] { R.id }},
                    { A.relIds,           new [] { R.cs, R.dm, R.lo, R.qs }},
                    //{ a14:imgLayer,     new [] { R.embed }},
                    //{ ax:ocx,           new [] { R.id }},
                    { C.chart,            new [] { R.id }},
                    { C.externalData,     new [] { R.id }},
                    { C.userShapes,       new [] { R.id }},
                    { DGM.relIds,         new [] { R.cs, R.dm, R.lo, R.qs }},
                    { O.OLEObject,        new [] { R.id }},
                    { VML.fill,           new [] { R.id }},
                    { VML.imagedata,      new [] { R.href, R.id, R.pict }},
                    { VML.stroke,         new [] { R.id }},
                    { W.altChunk,         new [] { R.id }},
                    { W.attachedTemplate, new [] { R.id }},
                    { W.control,          new [] { R.id }},
                    { W.dataSource,       new [] { R.id }},
                    { W.embedBold,        new [] { R.id }},
                    { W.embedBoldItalic,  new [] { R.id }},
                    { W.embedItalic,      new [] { R.id }},
                    { W.embedRegular,     new [] { R.id }},
                    { W.footerReference,  new [] { R.id }},
                    { W.headerReference,  new [] { R.id }},
                    { W.headerSource,     new [] { R.id }},
                    { W.hyperlink,        new [] { R.id }},
                    { W.printerSettings,  new [] { R.id }},
                    { W.recipientData,    new [] { R.id }},  // Mail merge, not required
                    { W.saveThroughXslt,  new [] { R.id }},
                    { W.sourceFileName,   new [] { R.id }},  // Framesets, not required
                    { W.src,              new [] { R.id }},  // Mail merge, not required
                    { W.subDoc,           new [] { R.id }},  // Sub documents, not required
                    //{ w14:contentPart,  new [] { R.id }},
                    { WNE.toolbarData,    new [] { R.id }},
                };
        }

        private static void CopySpecifiedCustomXmlParts(WordprocessingDocument sourceDocument, WordprocessingDocument output, DocumentBuilderSettings settings)
        {
            if (settings.CustomXmlGuidList == null || !settings.CustomXmlGuidList.Any())
                return;

            foreach (var customXmlPart in sourceDocument.MainDocumentPart.CustomXmlParts)
            {
                var propertyPart = customXmlPart
                    .Parts
                    .Select(p => p.OpenXmlPart)
                    .Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
                    .FirstOrDefault();
                if (propertyPart != null)
                {
                    var propertyPartDoc = propertyPart.GetXDocument();
#if false
        At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
        querying or generating that markup.  This is an example of the Custom XML Properties part.

<ds:datastoreItem ds:itemID="{1337A0C2-E6EE-4612-ACA5-E0E5A513381D}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml">
  <ds:schemaRefs />
</ds:datastoreItem>
#endif
                    var itemID = (string)propertyPartDoc.Root.Attribute(DS.itemID);
                    if (itemID != null)
                    {
                        itemID = itemID.Trim('{', '}');
                        if (settings.CustomXmlGuidList.Contains(itemID))
                        {
                            var newPart = output.MainDocumentPart.AddCustomXmlPart(customXmlPart.ContentType);
                            newPart.GetXDocument().Add(customXmlPart.GetXDocument().Root);
                            foreach (var propPart in customXmlPart.Parts.Select(p => p.OpenXmlPart))
                            {
                                var newPropPart = newPart.AddNewPart<CustomXmlPropertiesPart>();
                                newPropPart.GetXDocument().Add(propPart.GetXDocument().Root);
                            }
                        }
                    }
                }
            }
        }

        private static void RemoveHeadersAndFootersFromSections(WordprocessingDocument doc)
        {
            var mdXDoc = doc.MainDocumentPart.GetXDocument();
            var sections = mdXDoc.Descendants(W.sectPr).ToList();
            foreach (var sect in sections)
            {
                sect.Elements(W.headerReference).Remove();
                sect.Elements(W.footerReference).Remove();
            }
            doc.MainDocumentPart.PutXDocument();
        }

        private class CachedHeaderFooter
        {
            public XName Ref;
            public string Type;
            public string CachedPartRid;
        };

        private static void ProcessSectionsForLinkToPreviousHeadersAndFooters(WordprocessingDocument doc)
        {
            var cachedHeaderFooter = new[]
            {
                new CachedHeaderFooter { Ref = W.headerReference, Type = "first" },
                new CachedHeaderFooter { Ref = W.headerReference, Type = "even" },
                new CachedHeaderFooter { Ref = W.headerReference, Type = "default" },
                new CachedHeaderFooter { Ref = W.footerReference, Type = "first" },
                new CachedHeaderFooter { Ref = W.footerReference, Type = "even" },
                new CachedHeaderFooter { Ref = W.footerReference, Type = "default" },
            };

            var mdXDoc = doc.MainDocumentPart.GetXDocument();
            var sections = mdXDoc.Descendants(W.sectPr).ToList();
            var firstSection = true;
            foreach (var sect in sections)
            {
                if (firstSection)
                {
                    var headerFirst = FindReference(sect, W.headerReference, "first");
                    var headerDefault = FindReference(sect, W.headerReference, "default");
                    var headerEven = FindReference(sect, W.headerReference, "even");
                    var footerFirst = FindReference(sect, W.footerReference, "first");
                    var footerDefault = FindReference(sect, W.footerReference, "default");
                    var footerEven = FindReference(sect, W.footerReference, "even");

                    if (headerEven == null)
                    {
                        if (headerDefault != null)
                            AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, (string)headerDefault.Attribute(R.id), W.headerReference, "even");
                        else
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, W.headerReference, "even");
                    }

                    if (headerFirst == null)
                    {
                        if (headerDefault != null)
                            AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, (string)headerDefault.Attribute(R.id), W.headerReference, "first");
                        else
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, W.headerReference, "first");
                    }

                    if (footerEven == null)
                    {
                        if (footerDefault != null)
                            AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, (string)footerDefault.Attribute(R.id), W.footerReference, "even");
                        else
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, W.footerReference, "even");
                    }

                    if (footerFirst == null)
                    {
                        if (footerDefault != null)
                            AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, (string)footerDefault.Attribute(R.id), W.footerReference, "first");
                        else
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, W.footerReference, "first");
                    }

                    foreach (var hf in cachedHeaderFooter)
                    {
                        if (sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type) == null)
                            InitEmptyHeaderOrFooter(doc.MainDocumentPart, sect, hf.Ref, hf.Type);
                        var reference = sect.Elements(hf.Ref).FirstOrDefault(z => (string)z.Attribute(W.type) == hf.Type);
                        if (reference == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        hf.CachedPartRid = (string)reference.Attribute(R.id);
                    }
                    firstSection = false;
                    continue;
                }

                CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "first");
                CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "even");
                CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.headerReference, "default");
                CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "first");
                CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "even");
                CopyOrCacheHeaderOrFooter(doc, cachedHeaderFooter, sect, W.footerReference, "default");
            }
            doc.MainDocumentPart.PutXDocument();
        }

        private static void CopyOrCacheHeaderOrFooter(WordprocessingDocument doc, CachedHeaderFooter[] cachedHeaderFooter, XElement sect, XName referenceXName, string type)
        {
            var referenceElement = FindReference(sect, referenceXName, type);
            if (referenceElement == null)
            {
                var cachedPartRid = cachedHeaderFooter.FirstOrDefault(z => z.Ref == referenceXName && z.Type == type).CachedPartRid;
                AddReferenceToExistingHeaderOrFooter(doc.MainDocumentPart, sect, cachedPartRid, referenceXName, type);
            }
            else
            {
                var cachedPart = cachedHeaderFooter.FirstOrDefault(z => z.Ref == referenceXName && z.Type == type);
                cachedPart.CachedPartRid = (string)referenceElement.Attribute(R.id);
            }
        }

        private static XElement FindReference(XElement sect, XName reference, string type) =>
            sect.Elements(reference).FirstOrDefault(z => (string)z.Attribute(W.type) == type);

        private static void AddReferenceToExistingHeaderOrFooter(MainDocumentPart mainDocPart, XElement sect, string rId, XName reference, string toType)
        {
            if (reference == W.headerReference)
            {
                var referenceToAdd = new XElement(W.headerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, rId));
                sect.AddFirst(referenceToAdd);
            }
            else
            {
                var referenceToAdd = new XElement(W.footerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, rId));
                sect.AddFirst(referenceToAdd);
            }
        }

        private static void InitEmptyHeaderOrFooter(MainDocumentPart mainDocPart, XElement sect, XName referenceXName, string toType)
        {
            XDocument xDoc = null;
            if (referenceXName == W.headerReference)
            {
                xDoc = XDocument.Parse(
                    @"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                    <w:hdr xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
                           xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
                           xmlns:o='urn:schemas-microsoft-com:office:office'
                           xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                           xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
                           xmlns:v='urn:schemas-microsoft-com:vml'
                           xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
                           xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                           xmlns:w10='urn:schemas-microsoft-com:office:word'
                           xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                           xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
                           xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml'
                           xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
                           xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
                           xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
                           xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
                           mc:Ignorable='w14 w15 wp14'>
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val='Header' />
                        </w:pPr>
                        <w:r>
                          <w:t></w:t>
                        </w:r>
                      </w:p>
                    </w:hdr>");
                var newHeaderPart = mainDocPart.AddNewPart<HeaderPart>();
                newHeaderPart.PutXDocument(xDoc);
                var referenceToAdd = new XElement(W.headerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, mainDocPart.GetIdOfPart(newHeaderPart)));
                sect.AddFirst(referenceToAdd);
            }
            else
            {
                xDoc = XDocument.Parse(@"<?xml version='1.0' encoding='utf-8' standalone='yes'?>
                    <w:ftr xmlns:wpc='http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas'
                           xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006'
                           xmlns:o='urn:schemas-microsoft-com:office:office'
                           xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
                           xmlns:m='http://schemas.openxmlformats.org/officeDocument/2006/math'
                           xmlns:v='urn:schemas-microsoft-com:vml'
                           xmlns:wp14='http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
                           xmlns:wp='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
                           xmlns:w10='urn:schemas-microsoft-com:office:word'
                           xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                           xmlns:w14='http://schemas.microsoft.com/office/word/2010/wordml'
                           xmlns:w15='http://schemas.microsoft.com/office/word/2012/wordml'
                           xmlns:wpg='http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'
                           xmlns:wpi='http://schemas.microsoft.com/office/word/2010/wordprocessingInk'
                           xmlns:wne='http://schemas.microsoft.com/office/word/2006/wordml'
                           xmlns:wps='http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
                           mc:Ignorable='w14 w15 wp14'>
                      <w:p>
                        <w:pPr>
                          <w:pStyle w:val='Footer' />
                        </w:pPr>
                        <w:r>
                          <w:t></w:t>
                        </w:r>
                      </w:p>
                    </w:ftr>");
                var newFooterPart = mainDocPart.AddNewPart<FooterPart>();
                newFooterPart.PutXDocument(xDoc);
                var referenceToAdd = new XElement(W.footerReference,
                    new XAttribute(W.type, toType),
                    new XAttribute(R.id, mainDocPart.GetIdOfPart(newFooterPart)));
                sect.AddFirst(referenceToAdd);
            }
        }

        private static void TestPartForUnsupportedContent(OpenXmlPart part, int sourceNumber)
        {
            var obsoleteNamespaces = new[]
                {
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2007/5/30/wordml"),
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2008/9/16/wordprocessingDrawing"),
                    XNamespace.Get("http://schemas.microsoft.com/office/word/2009/2/wordml"),
                };
            var xDoc = part.GetXDocument();
            var invalidElement = xDoc.Descendants()
                .FirstOrDefault(d =>
                    {
                        var b = d.Name == W.subDoc ||
                                d.Name == W.control ||
                                d.Name == W.altChunk ||
                                d.Name.LocalName == "contentPart" ||
                                obsoleteNamespaces.Contains(d.Name.Namespace);
                        var b2 = b ||
                                 d.Attributes().Any(a => obsoleteNamespaces.Contains(a.Name.Namespace));
                        return b2;
                    });
            if (invalidElement != null)
            {
                if (invalidElement.Name == W.subDoc)
                    throw new DocumentBuilderException(
                        $"Source {sourceNumber} is unsupported document - contains sub document");
                if (invalidElement.Name == W.control)
                    throw new DocumentBuilderException(
                        $"Source {sourceNumber} is unsupported document - contains ActiveX controls");
                if (invalidElement.Name == W.altChunk)
                    throw new DocumentBuilderException(
                        $"Source {sourceNumber} is unsupported document - contains altChunk");
                if (invalidElement.Name.LocalName == "contentPart")
                    throw new DocumentBuilderException(
                        $"Source {sourceNumber} is unsupported document - contains contentPart content");
                if (obsoleteNamespaces.Contains(invalidElement.Name.Namespace) ||
                    invalidElement.Attributes().Any(a => obsoleteNamespaces.Contains(a.Name.Namespace)))
                    throw new DocumentBuilderException(
                        $"Source {sourceNumber} is unsupported document - contains obsolete namespace");
            }
        }

        //What does not work:
        //- sub docs
        //- bidi text appears to work but has not been tested
        //- languages other than en-us appear to work but have not been tested
        //- documents with activex controls
        //- mail merge source documents (look for dataSource in settings)
        //- documents with ink
        //- documents with frame sets and frames
        private static void TestForUnsupportedDocument(WordprocessingDocument doc, int sourceNumber)
        {
            if (doc.MainDocumentPart.GetXDocument().Root == null)
                throw new DocumentBuilderException(
                    $"Source {sourceNumber} is an invalid document - MainDocumentPart contains no content.");

            if (doc.MainDocumentPart.GetXDocument().Root.Name.NamespaceName == "http://purl.oclc.org/ooxml/wordprocessingml/main")
                throw new DocumentBuilderException($"Source {sourceNumber} is saved in strict mode, not supported");

            // note: if ever want to support section changes, need to address the code that rationalizes headers and footers, propagating to sections that inherit headers/footers from prev section
            foreach (var d in doc.MainDocumentPart.GetXDocument().Descendants())
            {
                if (d.Name == W.sectPrChange)
                    throw new DocumentBuilderException(
                        $"Source {sourceNumber} contains section changes (w:sectPrChange), not supported");

                // note: if ever want to support Open-Xml-PowerTools attributes, need to make sure that all attributes are propagated in all cases
                //if (d.Name.Namespace == PtOpenXml.ptOpenXml ||
                //    d.Name.Namespace == PtOpenXml.pt)
                //    throw new DocumentBuilderException(string.Format("Source {0} contains Open-Xml-PowerTools markup, not supported", sourceNumber));
                //if (d.Attributes().Any(a => a.Name.Namespace == PtOpenXml.ptOpenXml || a.Name.Namespace == PtOpenXml.pt))
                //    throw new DocumentBuilderException(string.Format("Source {0} contains Open-Xml-PowerTools markup, not supported", sourceNumber));
            }

            TestPartForUnsupportedContent(doc.MainDocumentPart, sourceNumber);
            foreach (var hdr in doc.MainDocumentPart.HeaderParts)
                TestPartForUnsupportedContent(hdr, sourceNumber);
            foreach (var ftr in doc.MainDocumentPart.FooterParts)
                TestPartForUnsupportedContent(ftr, sourceNumber);
            if (doc.MainDocumentPart.FootnotesPart != null)
                TestPartForUnsupportedContent(doc.MainDocumentPart.FootnotesPart, sourceNumber);
            if (doc.MainDocumentPart.EndnotesPart != null)
                TestPartForUnsupportedContent(doc.MainDocumentPart.EndnotesPart, sourceNumber);

            if (doc.MainDocumentPart.DocumentSettingsPart != null &&
                doc.MainDocumentPart.DocumentSettingsPart.GetXDocument().Descendants().Any(d => d.Name == W.src ||
                d.Name == W.recipientData || d.Name == W.mailMerge))
                throw new DocumentBuilderException(
                    $"Source {sourceNumber} is unsupported document - contains Mail Merge content");
            if (doc.MainDocumentPart.WebSettingsPart != null &&
                doc.MainDocumentPart.WebSettingsPart.GetXDocument().Descendants().Any(d => d.Name == W.frameset))
                throw new DocumentBuilderException(
                    $"Source {sourceNumber} is unsupported document - contains a frameset");
            var numberingElements = doc.MainDocumentPart
                .GetXDocument()
                .Descendants(W.numPr)
                .Where(n =>
                    {
                        var zeroId = (int?)n.Attribute(W.id) == 0;
                        var hasChildInsId = n.Elements(W.ins).Any();
                        if (zeroId || hasChildInsId)
                            return false;
                        return true;
                    })
                .ToList();
            if (numberingElements.Any() &&
                doc.MainDocumentPart.NumberingDefinitionsPart == null)
                throw new DocumentBuilderException(
                    $"Source {sourceNumber} is invalid document - contains numbering markup but no numbering part");
        }

        private static void FixUpSectionProperties(WordprocessingDocument newDocument)
        {
            var mainDocumentXDoc = newDocument.MainDocumentPart.GetXDocument();
            mainDocumentXDoc.Declaration.Standalone = Yes;
            mainDocumentXDoc.Declaration.Encoding = Utf8;
            var body = mainDocumentXDoc.Root.Element(W.body);
            var sectionPropertiesToMove = body
                .Elements()
                .Take(body.Elements().Count() - 1)
                .Where(e => e.Name == W.sectPr)
                .ToList();
            foreach (var s in sectionPropertiesToMove)
            {
                var p = s.SiblingsBeforeSelfReverseDocumentOrder().First();
                if (p.Element(W.pPr) == null)
                    p.AddFirst(new XElement(W.pPr));
                p.Element(W.pPr).Add(s);
            }
            foreach (var s in sectionPropertiesToMove)
                s.Remove();
        }

        private static void AddSectionAndDependencies(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            XElement sectionMarkup, List<ImageData> images)
        {
            var headerReferences = sectionMarkup.Elements(W.headerReference);
            foreach (var headerReference in headerReferences)
            {
                var oldRid = headerReference.Attribute(R.id).Value;
                HeaderPart oldHeaderPart = null;
                try
                {
                    oldHeaderPart = (HeaderPart)sourceDocument.MainDocumentPart.GetPartById(oldRid);
                }
                catch (ArgumentOutOfRangeException)
                {
                    var message = $"ArgumentOutOfRangeException, attempting to get header rId={oldRid}";
                    throw new OpenXmlPowerToolsException(message);
                }
                var oldHeaderXDoc = oldHeaderPart.GetXDocument();
                if (oldHeaderXDoc != null && oldHeaderXDoc.Root != null)
                    CopyNumbering(sourceDocument, newDocument, new[] { oldHeaderXDoc.Root }, images);
                var newHeaderPart = newDocument.MainDocumentPart.AddNewPart<HeaderPart>();
                var newHeaderXDoc = newHeaderPart.GetXDocument();
                newHeaderXDoc.Declaration.Standalone = Yes;
                newHeaderXDoc.Declaration.Encoding = Utf8;
                newHeaderXDoc.Add(oldHeaderXDoc.Root);
                headerReference.SetAttributeValue(R.id, newDocument.MainDocumentPart.GetIdOfPart(newHeaderPart));
                AddRelationships(oldHeaderPart, newHeaderPart, new[] { newHeaderXDoc.Root });
                CopyRelatedPartsForContentParts(oldHeaderPart, newHeaderPart, new[] { newHeaderXDoc.Root }, images);
            }

            var footerReferences = sectionMarkup.Elements(W.footerReference);
            foreach (var footerReference in footerReferences)
            {
                var oldRid = footerReference.Attribute(R.id).Value;
                var oldFooterPart2 = sourceDocument.MainDocumentPart.GetPartById(oldRid);
                if (!(oldFooterPart2 is FooterPart))
                    throw new DocumentBuilderException("Invalid document - invalid footer part.");

                var oldFooterPart = (FooterPart)oldFooterPart2;
                var oldFooterXDoc = oldFooterPart.GetXDocument();
                if (oldFooterXDoc != null && oldFooterXDoc.Root != null)
                    CopyNumbering(sourceDocument, newDocument, new[] { oldFooterXDoc.Root }, images);
                var newFooterPart = newDocument.MainDocumentPart.AddNewPart<FooterPart>();
                var newFooterXDoc = newFooterPart.GetXDocument();
                newFooterXDoc.Declaration.Standalone = Yes;
                newFooterXDoc.Declaration.Encoding = Utf8;
                newFooterXDoc.Add(oldFooterXDoc.Root);
                footerReference.SetAttributeValue(R.id, newDocument.MainDocumentPart.GetIdOfPart(newFooterPart));
                AddRelationships(oldFooterPart, newFooterPart, new[] { newFooterXDoc.Root });
                CopyRelatedPartsForContentParts(oldFooterPart, newFooterPart, new[] { newFooterXDoc.Root }, images);
            }
        }

        private static void MergeStyles(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, XDocument fromStyles, XDocument toStyles, IEnumerable<XElement> newContent)
        {
#if MergeStylesWithSameNames
            var newIds = new Dictionary<string, string>();
#endif
            if (fromStyles.Root == null)
                return;

            foreach (var style in fromStyles.Root.Elements(W.style))
            {
                var fromId = (string)style.Attribute(W.styleId);
                var fromName = (string)style.Elements(W.name).Attributes(W.val).FirstOrDefault();

                var toStyle = toStyles
                    .Root
                    .Elements(W.style)
                    .FirstOrDefault(st => (string)st.Elements(W.name).Attributes(W.val).FirstOrDefault() == fromName);

                if (toStyle == null)
                {
#if MergeStylesWithSameNames
                    var linkElement = style.Element(W.link);
                    if (linkElement != null && newIds.TryGetValue(linkElement.Attribute(W.val).Value, out var linkedId))
                    {
                        var linkedStyle = toStyles.Root.Elements(W.style)
                            .First(o => o.Attribute(W.styleId).Value == linkedId);
                        if (linkedStyle.Element(W.link) != null)
                            newIds.Add(fromId, linkedStyle.Element(W.link).Attribute(W.val).Value);
                        continue;
                    }

                    //string name = (string)style.Elements(W.name).Attributes(W.val).FirstOrDefault();
                    //var namedStyle = toStyles
                    //    .Root
                    //    .Elements(W.style)
                    //    .Where(st => st.Element(W.name) != null)
                    //    .FirstOrDefault(o => (string)o.Element(W.name).Attribute(W.val) == name);
                    //if (namedStyle != null)
                    //{
                    //    if (! newIds.ContainsKey(fromId))
                    //        newIds.Add(fromId, namedStyle.Attribute(W.styleId).Value);
                    //    continue;
                    //}
#endif

                    var number = 1;
                    var abstractNumber = 0;
                    XDocument oldNumbering = null;
                    XDocument newNumbering = null;
                    foreach (var numReference in style.Descendants(W.numPr))
                    {
                        var idElement = numReference.Descendants(W.numId).FirstOrDefault();
                        if (idElement != null)
                        {
                            if (oldNumbering == null)
                            {
                                if (sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                                    oldNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                else
                                {
                                    oldNumbering = new XDocument();
                                    oldNumbering.Declaration = new XDeclaration(OnePointZero, Utf8, Yes);
                                    oldNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                                }
                            }
                            if (newNumbering == null)
                            {
                                if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                                {
                                    newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                    newNumbering.Declaration.Standalone = Yes;
                                    newNumbering.Declaration.Encoding = Utf8;
                                    var numIds = newNumbering
                                        .Root
                                        .Elements(W.num)
                                        .Select(f => (int)f.Attribute(W.numId));
                                    if (numIds.Any())
                                        number = numIds.Max() + 1;
                                    numIds = newNumbering
                                        .Root
                                        .Elements(W.abstractNum)
                                        .Select(f => (int)f.Attribute(W.abstractNumId));
                                    if (numIds.Any())
                                        abstractNumber = numIds.Max() + 1;
                                }
                                else
                                {
                                    newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                                    newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                                    newNumbering.Declaration.Standalone = Yes;
                                    newNumbering.Declaration.Encoding = Utf8;
                                    newNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                                }
                            }
                            var numId = idElement.Attribute(W.val).Value;
                            if (numId != "0")
                            {
                                var element = oldNumbering
                                    .Descendants()
                                    .Elements(W.num)
                                    .Where(p => ((string)p.Attribute(W.numId)) == numId)
                                    .FirstOrDefault();

                                // Copy abstract numbering element, if necessary (use matching NSID)
                                var abstractNumId = string.Empty;
                                if (element != null)
                                {
                                    abstractNumId = element
                                       .Elements(W.abstractNumId)
                                       .First()
                                       .Attribute(W.val)
                                       .Value;

                                    var abstractElement = oldNumbering
                                        .Descendants()
                                        .Elements(W.abstractNum)
                                        .Where(p => ((string)p.Attribute(W.abstractNumId)) == abstractNumId)
                                        .FirstOrDefault();
                                    var abstractNSID = string.Empty;
                                    if (abstractElement != null)
                                    {
                                        var nsidElement = abstractElement
                                            .Element(W.nsid);
                                        abstractNSID = null;
                                        if (nsidElement != null)
                                            abstractNSID = (string)nsidElement
                                                .Attribute(W.val);

                                        var newAbstractElement = newNumbering
                                            .Descendants()
                                            .Elements(W.abstractNum)
                                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                            .Where(p =>
                                            {
                                                var thisNsidElement = p.Element(W.nsid);
                                                if (thisNsidElement == null)
                                                    return false;
                                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                                            })
                                            .FirstOrDefault();
                                        if (newAbstractElement == null)
                                        {
                                            newAbstractElement = new XElement(abstractElement);
                                            newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                                            abstractNumber++;
                                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                                            else
                                                newNumbering.Root.Add(newAbstractElement);

                                            foreach (var pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                                            {
                                                var bulletId = (string)pictId.Attribute(W.val);
                                                var numPicBullet = oldNumbering
                                                    .Descendants(W.numPicBullet)
                                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                                var maxNumPicBulletId = new int[] { -1 }.Concat(
                                                    newNumbering.Descendants(W.numPicBullet)
                                                    .Attributes(W.numPicBulletId)
                                                    .Select(a => (int)a))
                                                    .Max() + 1;
                                                var newNumPicBullet = new XElement(numPicBullet);
                                                newNumPicBullet.Attribute(W.numPicBulletId).Value = maxNumPicBulletId.ToString();
                                                pictId.Attribute(W.val).Value = maxNumPicBulletId.ToString();
                                                newNumbering.Root.AddFirst(newNumPicBullet);
                                            }
                                        }
                                        var newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                                        // Copy numbering element, if necessary (use matching element with no overrides)
                                        XElement newElement = null;
                                        if (!element.Elements(W.lvlOverride).Any())
                                            newElement = newNumbering
                                                .Descendants()
                                                .Elements(W.num)
                                                .Where(p => !p.Elements(W.lvlOverride).Any() &&
                                                    ((string)p.Elements(W.abstractNumId).First().Attribute(W.val)) == newAbstractId)
                                                .FirstOrDefault();
                                        if (newElement == null)
                                        {
                                            newElement = new XElement(element);
                                            newElement
                                                .Elements(W.abstractNumId)
                                                .First()
                                                .Attribute(W.val).Value = newAbstractId;
                                            newElement.Attribute(W.numId).Value = number.ToString();
                                            number++;
                                            newNumbering.Root.Add(newElement);
                                        }
                                        idElement.Attribute(W.val).Value = newElement.Attribute(W.numId).Value;
                                    }
                                }
                            }
                        }
                    }

                    var newStyle = new XElement(style);
                    // get rid of anything not in the w: namespace
                    newStyle.Descendants().Where(d => d.Name.NamespaceName != W.w).Remove();
                    newStyle.Descendants().Attributes().Where(d => d.Name.NamespaceName != W.w).Remove();
                    toStyles.Root.Add(newStyle);
                }
                else
                {
                    var toId = (string)toStyle.Attribute(W.styleId);
                    if (fromId != toId)
                    {
                        if (! newIds.ContainsKey(fromId))
                            newIds.Add(fromId, toId);
                    }
                }
            }

#if MergeStylesWithSameNames
            if (newIds.Count > 0)
            {
                foreach (var style in toStyles
                    .Root
                    .Elements(W.style))
                {
                    ConvertToNewId(style.Element(W.basedOn), newIds);
                    ConvertToNewId(style.Element(W.next), newIds);
                }

                foreach (var item in newContent.DescendantsAndSelf()
                    .Where(d => d.Name == W.pStyle ||
                                d.Name == W.rStyle ||
                                d.Name == W.tblStyle))
                {
                    ConvertToNewId(item, newIds);
                }

                if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                {
                    var newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    ConvertNumberingPartToNewIds(newNumbering, newIds);
                }

                // Convert source document, since numberings will be copied over after styles.
                if (sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                {
                    var sourceNumbering = sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    ConvertNumberingPartToNewIds(sourceNumbering, newIds);
                }
            }
#endif
        }

        private static void MergeLatentStyles(XDocument fromStyles, XDocument toStyles)
        {
            var fromLatentStyles = fromStyles.Descendants(W.latentStyles).FirstOrDefault();
            if (fromLatentStyles == null)
                return;

            var toLatentStyles = toStyles.Descendants(W.latentStyles).FirstOrDefault();
            if (toLatentStyles == null)
            {
                var newLatentStylesElement = new XElement(W.latentStyles,
                    fromLatentStyles.Attributes());
                var globalDefaults = toStyles
                    .Descendants(W.docDefaults)
                    .FirstOrDefault();
                if (globalDefaults == null)
                {
                    var firstStyle = toStyles
                        .Root
                        .Elements(W.style)
                        .FirstOrDefault();
                    if (firstStyle == null)
                        toStyles.Root.Add(newLatentStylesElement);
                    else
                        firstStyle.AddBeforeSelf(newLatentStylesElement);
                }
                else
                    globalDefaults.AddAfterSelf(newLatentStylesElement);
            }
            toLatentStyles = toStyles.Descendants(W.latentStyles).FirstOrDefault();
            if (toLatentStyles == null)
                throw new OpenXmlPowerToolsException("Internal error");

            var toStylesHash = new HashSet<string>();
            foreach (var lse in toLatentStyles.Elements(W.lsdException))
                toStylesHash.Add((string)lse.Attribute(W.name));

            foreach (var fls in fromLatentStyles.Elements(W.lsdException))
            {
                var name = (string)fls.Attribute(W.name);
                if (toStylesHash.Contains(name))
                    continue;
                toLatentStyles.Add(fls);
                toStylesHash.Add(name);
            }

            var count = toLatentStyles
                .Elements(W.lsdException)
                .Count();

            toLatentStyles.SetAttributeValue(W.count, count);
        }

        private static void MergeDocDefaultStyles(XDocument xDocument, XDocument newXDoc)
        {
            var docDefaultStyles = xDocument.Descendants(W.docDefaults);
            foreach (var docDefaultStyle in docDefaultStyles)
            {
                newXDoc.Root.Add(docDefaultStyle);
            }
        }

#if MergeStylesWithSameNames
        private static void ConvertToNewId(XElement element, Dictionary<string, string> newIds)
        {
            if (element == null)
                return;

            var valueAttribute = element.Attribute(W.val);
            if (newIds.TryGetValue(valueAttribute.Value, out var newId))
            {
                valueAttribute.Value = newId;
            }
        }

        private static void ConvertNumberingPartToNewIds(XDocument newNumbering, Dictionary<string, string> newIds)
        {
            foreach (var abstractNum in newNumbering
                .Root
                .Elements(W.abstractNum))
            {
                ConvertToNewId(abstractNum.Element(W.styleLink), newIds);
                ConvertToNewId(abstractNum.Element(W.numStyleLink), newIds);
            }

            foreach (var item in newNumbering
                .Descendants()
                .Where(d => d.Name == W.pStyle ||
                            d.Name == W.rStyle ||
                            d.Name == W.tblStyle))
            {
                ConvertToNewId(item, newIds);
            }
        }
#endif

        private static void MergeFontTables(XDocument fromFontTable, XDocument toFontTable)
        {
            foreach (var font in fromFontTable.Root.Elements(W.font))
            {
                var name = font.Attribute(W.name).Value;
                if (toFontTable
                    .Root
                    .Elements(W.font)
                    .Where(o => o.Attribute(W.name).Value == name)
                    .Count() == 0)
                    toFontTable.Root.Add(new XElement(font));
            }
        }

        private static void CopyStylesAndFonts(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.StyleDefinitionsPart != null)
            {
                var oldStyles = sourceDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StyleDefinitionsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    var newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    newStyles.Declaration.Standalone = Yes;
                    newStyles.Declaration.Encoding = Utf8;
                    newStyles.Add(oldStyles.Root);
                }
                else
                {
                    var newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    MergeLatentStyles(oldStyles, newStyles);
                }
            }

            // Copy all styles with effects to the new document
            if (sourceDocument.MainDocumentPart.StylesWithEffectsPart != null)
            {
                var oldStyles = sourceDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StylesWithEffectsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StylesWithEffectsPart>();
                    var newStyles = newDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                    newStyles.Declaration.Standalone = Yes;
                    newStyles.Declaration.Encoding = Utf8;
                    newStyles.Add(oldStyles.Root);
                }
                else
                {
                    var newStyles = newDocument.MainDocumentPart.StylesWithEffectsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    MergeLatentStyles(oldStyles, newStyles);
                }
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.FontTablePart != null)
            {
                var oldFontTable = sourceDocument.MainDocumentPart.FontTablePart.GetXDocument();
                if (newDocument.MainDocumentPart.FontTablePart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                    var newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    newFontTable.Declaration.Standalone = Yes;
                    newFontTable.Declaration.Encoding = Utf8;
                    newFontTable.Add(oldFontTable.Root);
                }
                else
                {
                    var newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    MergeFontTables(oldFontTable, newFontTable);
                }
            }
        }

        private static void CopyComments(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var commentIdMap = new Dictionary<int, int>();
            var number = 0;
            XDocument oldComments = null;
            XDocument newComments = null;
            foreach (var comment in newContent.DescendantsAndSelf(W.commentReference))
            {
                if (oldComments == null)
                    oldComments = sourceDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                if (newComments == null)
                {
                    if (newDocument.MainDocumentPart.WordprocessingCommentsPart != null)
                    {
                        newComments = newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.Standalone = Yes;
                        newComments.Declaration.Encoding = Utf8;
                        var ids = newComments.Root.Elements(W.comment).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<WordprocessingCommentsPart>();
                        newComments = newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument();
                        newComments.Declaration.Standalone = Yes;
                        newComments.Declaration.Encoding = Utf8;
                        newComments.Add(new XElement(W.comments, NamespaceAttributes));
                    }
                }

                if (!int.TryParse((string)comment.Attribute(W.id), out var id))
                    throw new DocumentBuilderException("Invalid document - invalid comment id");
                var element = oldComments
                    .Descendants()
                    .Elements(W.comment)
                    .Where(p => {
                        if (! int.TryParse((string)p.Attribute(W.id), out var thisId))
                            throw new DocumentBuilderException("Invalid document - invalid comment id");
                        return thisId == id;
                    })
                    .FirstOrDefault();
                if (element == null)
                    throw new DocumentBuilderException("Invalid document - comment reference without associated comment in comments part");
                var newElement = new XElement(element);
                newElement.SetAttributeValue(W.id, number.ToString());
                newComments.Root.Add(newElement);
                if (! commentIdMap.ContainsKey(id))
                    commentIdMap.Add(id, number);
                number++;
            }
            foreach (var item in newContent.DescendantsAndSelf()
                .Where(d => d.Name == W.commentReference ||
                            d.Name == W.commentRangeStart ||
                            d.Name == W.commentRangeEnd)
                .ToList())
            {
                if (commentIdMap.ContainsKey((int)item.Attribute(W.id)))
                    item.SetAttributeValue(W.id, commentIdMap[(int)item.Attribute(W.id)].ToString());
            }
            if (sourceDocument.MainDocumentPart.WordprocessingCommentsPart != null &&
                newDocument.MainDocumentPart.WordprocessingCommentsPart != null)
            {
                AddRelationships(sourceDocument.MainDocumentPart.WordprocessingCommentsPart,
                    newDocument.MainDocumentPart.WordprocessingCommentsPart,
                    new[] { newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart.WordprocessingCommentsPart,
                    newDocument.MainDocumentPart.WordprocessingCommentsPart,
                    new[] { newDocument.MainDocumentPart.WordprocessingCommentsPart.GetXDocument().Root },
                    images);
            }
        }

        private static void AdjustUniqueIds(WordprocessingDocument sourceDocument,
            WordprocessingDocument newDocument, IEnumerable<XElement> newContent)
        {
            // adjust bookmark unique ids
            var maxId = 0;
            if (newDocument.MainDocumentPart.GetXDocument().Descendants(W.bookmarkStart).Any())
                maxId = newDocument.MainDocumentPart.GetXDocument().Descendants(W.bookmarkStart)
                    .Select(d => (int)d.Attribute(W.id)).Max();
            var bookmarkIdMap = new Dictionary<int, int>();
            foreach (var item in newContent.DescendantsAndSelf().Where(bm => bm.Name == W.bookmarkStart ||
                bm.Name == W.bookmarkEnd))
            {
                if (!int.TryParse((string)item.Attribute(W.id), out var id))
                    throw new DocumentBuilderException("Invalid document - invalid value for bookmark ID");
                if (!bookmarkIdMap.ContainsKey(id))
                    bookmarkIdMap.Add(id, ++maxId);
            }
            foreach (var bookmarkElement in newContent.DescendantsAndSelf().Where(e => e.Name == W.bookmarkStart ||
                e.Name == W.bookmarkEnd))
                bookmarkElement.SetAttributeValue(W.id, bookmarkIdMap[(int)bookmarkElement.Attribute(W.id)].ToString());

            // adjust shape unique ids
            // This doesn't work because OLEObjects refer to shapes by ID.
            // Punting on this, because sooner or later, this will be a non-issue.
            //foreach (var item in newContent.DescendantsAndSelf(VML.shape))
            //{
            //    Guid g = Guid.NewGuid();
            //    string s = "R" + g.ToString().Replace("-", "");
            //    item.Attribute(NoNamespace.id).Value = s;
            //}
        }

        private static void AdjustDocPrIds(WordprocessingDocument newDocument)
        {
            var docPrId = 0;
            foreach (var item in newDocument.MainDocumentPart.GetXDocument().Descendants(WP.docPr))
                item.SetAttributeValue(NoNamespace.id, (++docPrId).ToString());
            foreach (var header in newDocument.MainDocumentPart.HeaderParts)
                foreach (var item in header.GetXDocument().Descendants(WP.docPr))
                    item.SetAttributeValue(NoNamespace.id, (++docPrId).ToString());
            foreach (var footer in newDocument.MainDocumentPart.FooterParts)
                foreach (var item in footer.GetXDocument().Descendants(WP.docPr))
                    item.SetAttributeValue(NoNamespace.id, (++docPrId).ToString());
            if (newDocument.MainDocumentPart.FootnotesPart != null)
                foreach (var item in newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Descendants(WP.docPr))
                    item.SetAttributeValue(NoNamespace.id, (++docPrId).ToString());
            if (newDocument.MainDocumentPart.EndnotesPart != null)
                foreach (var item in newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Descendants(WP.docPr))
                    item.SetAttributeValue(NoNamespace.id, (++docPrId).ToString());
        }

        // This probably doesn't need to be done, except that the Open XML SDK will not validate
        // documents that contain the o:gfxdata attribute.
        private static void RemoveGfxdata(IEnumerable<XElement> newContent)
        {
            newContent.DescendantsAndSelf().Attributes(O.gfxdata).Remove();
        }

        private static object InsertTransform(XNode node, List<XElement> newContent)
        {
            if (!(node is XElement element))
                return node;

            if (element.Annotation<ReplaceSemaphore>() != null)
                return newContent;
            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => InsertTransform(n, newContent)));
        }

        private class ReplaceSemaphore { }

        // Rules for sections
        // - if KeepSections for all documents in the source collection are false, then it takes the section
        //   from the first document.
        // - if you specify true for any document, and if the last section is part of the specified content,
        //   then that section is copied.  If any paragraph in the content has a section, then that section
        //   is copied.
        // - if you specify true for any document, and there are no sections for any paragraphs, then no
        //   sections are copied.
        private static void AppendDocument(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            List<XElement> newContent, bool keepSection, string insertId, List<ImageData> images)
        {
            FixRanges(sourceDocument.MainDocumentPart.GetXDocument(), newContent);
            AddRelationships(sourceDocument.MainDocumentPart, newDocument.MainDocumentPart, newContent);
            CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart, newDocument.MainDocumentPart,
                newContent, images);

            // Append contents
            var newMainXDoc = newDocument.MainDocumentPart.GetXDocument();
            newMainXDoc.Declaration.Standalone = Yes;
            newMainXDoc.Declaration.Encoding = Utf8;
            if (keepSection == false)
            {
                var adjustedContents = newContent.Where(e => e.Name != W.sectPr).ToList();
                adjustedContents.DescendantsAndSelf(W.sectPr).Remove();
                newContent = adjustedContents;
            }
            var listOfSectionProps = newContent.DescendantsAndSelf(W.sectPr).ToList();
            foreach (var sectPr in listOfSectionProps)
                AddSectionAndDependencies(sourceDocument, newDocument, sectPr, images);
            CopyStylesAndFonts(sourceDocument, newDocument, newContent);
            CopyNumbering(sourceDocument, newDocument, newContent, images);
            CopyComments(sourceDocument, newDocument, newContent, images);
            CopyFootnotes(sourceDocument, newDocument, newContent, images);
            CopyEndnotes(sourceDocument, newDocument, newContent, images);
            AdjustUniqueIds(sourceDocument, newDocument, newContent);
            RemoveGfxdata(newContent);
            CopyCustomXmlPartsForDataBoundContentControls(sourceDocument, newDocument, newContent);
            CopyWebExtensions(sourceDocument, newDocument);
            if (insertId != null)
            {
                var insertElementToReplace = newMainXDoc
                    .Descendants(PtOpenXml.Insert)
                    .FirstOrDefault(i => (string)i.Attribute(PtOpenXml.Id) == insertId);
                insertElementToReplace?.AddAnnotation(new ReplaceSemaphore());
                newMainXDoc.Element(W.document).ReplaceWith((XElement)InsertTransform(newMainXDoc.Root, newContent));
            }
            else
                newMainXDoc.Root.Element(W.body).Add(newContent);

            if (newMainXDoc.Descendants().Any(d =>
            {
                if (d.Name.Namespace == PtOpenXml.pt || d.Name.Namespace == PtOpenXml.ptOpenXml)
                    return true;
                if (d.Attributes().Any(att => att.Name.Namespace == PtOpenXml.pt || att.Name.Namespace == PtOpenXml.ptOpenXml))
                    return true;
                return false;
            }))
            {
                var root = newMainXDoc.Root;
                if (!root.Attributes().Any(na => na.Value == PtOpenXml.pt.NamespaceName))
                {
                    root.Add(new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt.NamespaceName));
                    AddToIgnorable(root, "pt");
                }
                if (!root.Attributes().Any(na => na.Value == PtOpenXml.ptOpenXml.NamespaceName))
                {
                    root.Add(new XAttribute(XNamespace.Xmlns + "pt14", PtOpenXml.ptOpenXml.NamespaceName));
                    AddToIgnorable(root, "pt14");
                }
            }
        }

        private static void CopyWebExtensions(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument)
        {
            if (sourceDocument.WebExTaskpanesPart == null
                || newDocument.WebExTaskpanesPart != null)
                return;

            newDocument.AddWebExTaskpanesPart();
            newDocument.WebExTaskpanesPart.GetXDocument().Add(sourceDocument.WebExTaskpanesPart.GetXDocument().Root);

            foreach (var sourceWebExtensionPart in sourceDocument.WebExTaskpanesPart.WebExtensionParts)
            {
                var newWebExtensionPart = newDocument.WebExTaskpanesPart.AddNewPart<WebExtensionPart>(
                    sourceDocument.WebExTaskpanesPart.GetIdOfPart(sourceWebExtensionPart));
                newWebExtensionPart.GetXDocument().Add(sourceWebExtensionPart.GetXDocument().Root);
            }
        }

        private static void AddToIgnorable(XElement root, string v)
        {
            var ignorable = root.Attribute(MC.Ignorable);
            if (ignorable is null)
                return;

            var val = (string)ignorable;
            val = val + " " + v;
            ignorable.Remove();
            root.SetAttributeValue(MC.Ignorable, val);
        }

        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /// New method to support new functionality
        private static void AppendDocument(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, OpenXmlPart part,
            List<XElement> newContent, bool keepSection, string insertId, List<ImageData> images)
        {
            // Append contents
            var partXDoc = part.GetXDocument();
            partXDoc.Declaration.Standalone = Yes;
            partXDoc.Declaration.Encoding = Utf8;

            FixRanges(part.GetXDocument(), newContent);
            AddRelationships(sourceDocument.MainDocumentPart, part, newContent);
            CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart, part,
                newContent, images);

            // never keep sections for content to be inserted into a header/footer
            var adjustedContents = newContent.Where(e => e.Name != W.sectPr).ToList();
            adjustedContents.DescendantsAndSelf(W.sectPr).Remove();
            newContent = adjustedContents;

            CopyNumbering(sourceDocument, newDocument, newContent, images);
            CopyComments(sourceDocument, newDocument, newContent, images);
            AdjustUniqueIds(sourceDocument, newDocument, newContent);
            RemoveGfxdata(newContent);

            if (insertId == null)
                throw new OpenXmlPowerToolsException("Internal error");

            var insertElementToReplace = partXDoc
                .Descendants(PtOpenXml.Insert)
                .FirstOrDefault(i => (string)i.Attribute(PtOpenXml.Id) == insertId);
            insertElementToReplace?.AddAnnotation(new ReplaceSemaphore());
            partXDoc.Elements().First().ReplaceWith((XElement)InsertTransform(partXDoc.Root, newContent));
        }
        /// ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public static WmlDocument ExtractGlossaryDocument(WmlDocument wmlGlossaryDocument)
        {
            if (RelationshipMarkup is null)
                InitRelationshipMarkup();

            using var ms = new MemoryStream();
            ms.Write(wmlGlossaryDocument.DocumentByteArray, 0, wmlGlossaryDocument.DocumentByteArray.Length);
            using var wDoc = WordprocessingDocument.Open(ms, false);

            var fromXd = wDoc.MainDocumentPart.GlossaryDocumentPart?.GetXDocument();
            if (fromXd?.Root is null)
                return null;

            using var outMs = new MemoryStream();
            using (var outWDoc = WordprocessingDocument.Create(outMs, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            {
                var images = new List<ImageData>();

                var mdp = outWDoc.AddMainDocumentPart();
                var mdpXd = mdp.GetXDocument();
                var root = new XElement(W.document);
                if (mdpXd.Root == null)
                    mdpXd.Add(root);
                else
                    mdpXd.Root.ReplaceWith(root);
                root.Add(new XElement(W.body,
                    fromXd.Root.Elements(W.docParts)));
                mdp.PutXDocument();

                var newContent = fromXd.Root.Elements(W.docParts);
                CopyGlossaryDocumentPartsFromGD(wDoc, outWDoc, newContent, images);
                CopyRelatedPartsForContentParts(wDoc.MainDocumentPart.GlossaryDocumentPart, mdp, newContent, images);
            }
            return new WmlDocument("Glossary.docx", outMs.ToArray());
        }

        private static void CopyGlossaryDocumentPartsFromGD(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart != null)
            {
                var oldStyles = sourceDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.GetXDocument();
                if (newDocument.MainDocumentPart.StyleDefinitionsPart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                    var newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    newStyles.Declaration.Standalone = Yes;
                    newStyles.Declaration.Encoding = Utf8;
                    newStyles.Add(oldStyles.Root);
                    newDocument.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
                }
                else
                {
                    var newStyles = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                    MergeStyles(sourceDocument, newDocument, oldStyles, newStyles, newContent);
                    newDocument.MainDocumentPart.StyleDefinitionsPart.PutXDocument();
                }
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart != null)
            {
                var oldFontTable = sourceDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart.GetXDocument();
                if (newDocument.MainDocumentPart.FontTablePart == null)
                {
                    newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                    var newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    newFontTable.Declaration.Standalone = Yes;
                    newFontTable.Declaration.Encoding = Utf8;
                    newFontTable.Add(oldFontTable.Root);
                    newDocument.MainDocumentPart.FontTablePart.PutXDocument();
                }
                else
                {
                    var newFontTable = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                    MergeFontTables(oldFontTable, newFontTable);
                    newDocument.MainDocumentPart.FontTablePart.PutXDocument();
                }
            }

            var oldSettingsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                var newSettingsPart = newDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                var settingsXDoc = oldSettingsPart.GetXDocument();
                AddRelationships(oldSettingsPart, newSettingsPart, new[] { settingsXDoc.Root });
                //CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                //CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                var newXDoc = newDocument.MainDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                CopyRelatedPartsForContentParts(oldSettingsPart, newSettingsPart, new[] { newXDoc.Root }, images);
                newSettingsPart.PutXDocument(newXDoc);
            }

            var oldWebSettingsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                var newWebSettingsPart = newDocument.MainDocumentPart.AddNewPart<WebSettingsPart>();
                var settingsXDoc = oldWebSettingsPart.GetXDocument();
                AddRelationships(oldWebSettingsPart, newWebSettingsPart, new[] { settingsXDoc.Root });
                var newXDoc = newDocument.MainDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                newWebSettingsPart.PutXDocument(newXDoc);
            }

            var oldNumberingDefinitionsPart = sourceDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart;
            if (oldNumberingDefinitionsPart != null)
            {
                CopyNumberingForGlossaryDocumentPartFromGD(oldNumberingDefinitionsPart, newDocument, newContent, images);
            }
        }

        private static void CopyGlossaryDocumentPartsToGD(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            // Copy all styles to the new document
            if (sourceDocument.MainDocumentPart.StyleDefinitionsPart != null)
            {
                var oldStyles = sourceDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<StyleDefinitionsPart>();
                var newStyles = newDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.GetXDocument();
                newStyles.Declaration.Standalone = Yes;
                newStyles.Declaration.Encoding = Utf8;
                newStyles.Add(oldStyles.Root);
                newDocument.MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart.PutXDocument();
            }

            // Copy fontTable to the new document
            if (sourceDocument.MainDocumentPart.FontTablePart != null)
            {
                var oldFontTable = sourceDocument.MainDocumentPart.FontTablePart.GetXDocument();
                newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<FontTablePart>();
                var newFontTable = newDocument.MainDocumentPart.GlossaryDocumentPart.FontTablePart.GetXDocument();
                newFontTable.Declaration.Standalone = Yes;
                newFontTable.Declaration.Encoding = Utf8;
                newFontTable.Add(oldFontTable.Root);
                newDocument.MainDocumentPart.FontTablePart.PutXDocument();
            }

            var oldSettingsPart = sourceDocument.MainDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                var newSettingsPart = newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<DocumentSettingsPart>();
                var settingsXDoc = oldSettingsPart.GetXDocument();
                AddRelationships(oldSettingsPart, newSettingsPart, new[] { settingsXDoc.Root });
                //CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                //CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                var newXDoc = newDocument.MainDocumentPart.GlossaryDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                CopyRelatedPartsForContentParts(oldSettingsPart, newSettingsPart, new[] { newXDoc.Root }, images);
                newSettingsPart.PutXDocument(newXDoc);
            }

            var oldWebSettingsPart = sourceDocument.MainDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                var newWebSettingsPart = newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<WebSettingsPart>();
                var settingsXDoc = oldWebSettingsPart.GetXDocument();
                AddRelationships(oldWebSettingsPart, newWebSettingsPart, new[] { settingsXDoc.Root });
                var newXDoc = newDocument.MainDocumentPart.GlossaryDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                newWebSettingsPart.PutXDocument(newXDoc);
            }

            var oldNumberingDefinitionsPart = sourceDocument.MainDocumentPart.NumberingDefinitionsPart;
            if (oldNumberingDefinitionsPart != null)
            {
                CopyNumberingForGlossaryDocumentPartToGD(oldNumberingDefinitionsPart, newDocument, newContent, images);
            }
        }


#if false
        At various locations in Open-Xml-PowerTools, you will find examples of Open XML markup that is associated with code associated with
        querying or generating that markup.  This is an example of the GlossaryDocument part.

<w:glossaryDocument xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid wp14">
  <w:docParts>
    <w:docPart>
      <w:docPartPr>
        <w:name w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
        <w:category>
          <w:name w:val="General" />
          <w:gallery w:val="placeholder" />
        </w:category>
        <w:types>
          <w:type w:val="bbPlcHdr" />
        </w:types>
        <w:behaviors>
          <w:behavior w:val="content" />
        </w:behaviors>
        <w:guid w:val="{13882A71-B5B7-4421-ACBB-9B61C61B3034}" />
      </w:docPartPr>
      <w:docPartBody>
        <w:p w:rsidR="00004EEA" w:rsidRDefault="00AD57F5" w:rsidP="00AD57F5">
          <w:pPr>
            <w:pStyle w:val="CDE7B64C7BB446AE905B622B0A882EB6" />
          </w:pPr>
          <w:r w:rsidRPr="00FB619D">
            <w:rPr>
              <w:rStyle w:val="PlaceholderText" />
              <w:lang w:val="da-DK" />
            </w:rPr>
            <w:t>Produktnavn</w:t>
          </w:r>
          <w:r w:rsidRPr="007379EE">
            <w:rPr>
              <w:rStyle w:val="PlaceholderText" />
            </w:rPr>
            <w:t>.</w:t>
          </w:r>
        </w:p>
      </w:docPartBody>
    </w:docPart>
  </w:docParts>
</w:glossaryDocument>
#endif

        private static void CopyCustomXmlPartsForDataBoundContentControls(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument, IEnumerable<XElement> newContent)
        {
            var itemList = new List<string>();
            foreach (var itemId in newContent
                .Descendants(W.dataBinding)
                .Select(e => (string)e.Attribute(W.storeItemID)))
                if (!itemList.Contains(itemId))
                    itemList.Add(itemId);
            foreach (var customXmlPart in sourceDocument.MainDocumentPart.CustomXmlParts)
            {
                var propertyPart = customXmlPart
                    .Parts
                    .Select(p => p.OpenXmlPart)
                    .Where(p => p.ContentType == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml")
                    .FirstOrDefault();
                if (propertyPart != null)
                {
                    var propertyPartDoc = propertyPart.GetXDocument();
                    if (itemList.Contains(propertyPartDoc.Root.Attribute(DS.itemID).Value))
                    {
                        var newPart = newDocument.MainDocumentPart.AddCustomXmlPart(customXmlPart.ContentType);
                        newPart.GetXDocument().Add(customXmlPart.GetXDocument().Root);
                        foreach (var propPart in customXmlPart.Parts.Select(p => p.OpenXmlPart))
                        {
                            var newPropPart = newPart.AddNewPart<CustomXmlPropertiesPart>();
                            newPropPart.GetXDocument().Add(propPart.GetXDocument().Root);
                        }
                    }
                }
            }
        }

        private static Dictionary<XName, XName[]> RelationshipMarkup = null;

        private static void UpdateContent(IEnumerable<XElement> newContent, XName elementToModify, string oldRid, string newRid)
        {
            foreach (var attributeName in RelationshipMarkup[elementToModify])
            {
                var elementsToUpdate = newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid);
                foreach (var element in elementsToUpdate)
                    element.SetAttributeValue(attributeName, newRid);
            }
        }

        private static void AddRelationships(OpenXmlPart oldPart, OpenXmlPart newPart, IEnumerable<XElement> newContent)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => RelationshipMarkup.ContainsKey(d.Name) &&
                    d.Attributes().Any(a => RelationshipMarkup[d.Name].Contains(a.Name)));
            foreach (var e in relevantElements)
            {
                if (e.Name == W.hyperlink)
                {
                    var relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink != null)
                        continue;
                    var newRid = Relationships.GetNewRelationshipId();
                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink == null)
                        continue;
                    //throw new DocumentBuilderInternalException("Internal Error 0002");
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == W.attachedTemplate || e.Name == W.saveThroughXslt)
                {
                    var relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    var newRid = Relationships.GetNewRelationshipId();
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new DocumentBuilderInternalException("Source {0} is invalid document - hyperlink contains invalid references");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == A.hlinkClick || e.Name == A.hlinkHover || e.Name == A.hlinkMouseOver)
                {
                    var relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink != null)
                        continue;
                    var newRid = Relationships.GetNewRelationshipId();
                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink == null)
                        continue;
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == VML.imagedata)
                {
                    var relId = (string)e.Attribute(R.href);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    var newRid = Relationships.GetNewRelationshipId();
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        throw new DocumentBuilderInternalException("Internal Error 0006");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                if (e.Name == A.blip)
                {
                    // <a:blip r:embed="rId6" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
                    var relId = (string)e.Attribute(R.link);
                    //if (relId == null)
                    //    relId = (string)e.Attribute(R.embed);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    var tempExternalRelationship = newPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempExternalRelationship != null)
                        continue;
                    var newRid = Relationships.GetNewRelationshipId();
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel == null)
                        continue;
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
            }
        }

        private class FromPreviousSourceSemaphore { };

        private static void CopyNumbering(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var numIdMap = new Dictionary<int, int>();
            var number = 1;
            var abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (var numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                var idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    oldNumbering ??= sourceDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            var numIds = newNumbering
                                .Root
                                .Elements(W.num)
                                .Select(f => (int)f.Attribute(W.numId));
                            if (numIds.Any())
                                number = numIds.Max() + 1;
                            numIds = newNumbering
                                .Root
                                .Elements(W.abstractNum)
                                .Select(f => (int)f.Attribute(W.abstractNumId));
                            if (numIds.Any())
                                abstractNumber = numIds.Max() + 1;
                        }
                        else
                        {
                            newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            newNumbering.Declaration.Standalone = Yes;
                            newNumbering.Declaration.Encoding = Utf8;
                            newNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                        }
                    }
                    var numId = (int)idElement.Attribute(W.val);
                    if (numId != 0)
                    {
                        var element = oldNumbering
                            .Descendants(W.num)
                            .Where(p => ((int)p.Attribute(W.numId)) == numId)
                            .FirstOrDefault();
                        if (element == null)
                            continue;

                        // Copy abstract numbering element, if necessary (use matching NSID)
                        var abstractNumIdStr = (string)element
                            .Elements(W.abstractNumId)
                            .First()
                            .Attribute(W.val);
                        if (!int.TryParse(abstractNumIdStr, out var abstractNumId))
                            throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");

                        var abstractElement = oldNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                            .First();
                        var nsidElement = abstractElement
                            .Element(W.nsid);
                        string abstractNSID = null;
                        if (nsidElement != null)
                            abstractNSID = (string)nsidElement
                                .Attribute(W.val);
                        var newAbstractElement = newNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                            .Where(p =>
                            {
                                var thisNsidElement = p.Element(W.nsid);
                                if (thisNsidElement == null)
                                    return false;
                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                            })
                            .FirstOrDefault();
                        if (newAbstractElement == null)
                        {
                            newAbstractElement = new XElement(abstractElement);
                            newAbstractElement.Attribute(W.abstractNumId).Value = abstractNumber.ToString();
                            abstractNumber++;
                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                            else
                                newNumbering.Root.Add(newAbstractElement);

                            foreach (var pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                            {
                                var bulletId = (string)pictId.Attribute(W.val);
                                var numPicBullet = oldNumbering
                                    .Descendants(W.numPicBullet)
                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                var maxNumPicBulletId = new int[] { -1 }.Concat(
                                    newNumbering.Descendants(W.numPicBullet)
                                    .Attributes(W.numPicBulletId)
                                    .Select(a => (int)a))
                                    .Max() + 1;
                                var newNumPicBullet = new XElement(numPicBullet);
                                newNumPicBullet.SetAttributeValue(W.numPicBulletId, maxNumPicBulletId.ToString());
                                pictId.SetAttributeValue(W.val, maxNumPicBulletId.ToString());
                                newNumbering.Root.AddFirst(newNumPicBullet);
                            }
                        }
                        var newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                        // Copy numbering element, if necessary (use matching element with no overrides)
                        XElement newElement;
                        if (numIdMap.ContainsKey(numId))
                        {
                            newElement = newNumbering
                                .Descendants()
                                .Elements(W.num)
                                .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                .First();
                        }
                        else
                        {
                            newElement = new XElement(element);
                            newElement
                                .Elements(W.abstractNumId)
                                .First()
                                .Attribute(W.val).Value = newAbstractId;
                            newElement.Attribute(W.numId).Value = number.ToString();
                            numIdMap.Add(numId, number);
                            number++;
                            newNumbering.Root.Add(newElement);
                        }
                        idElement.SetAttributeValue(W.val, newElement.Attribute(W.numId).Value);
                    }
                }
            }
            if (newNumbering != null)
            {
                foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                    abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                foreach (var num in newNumbering.Descendants(W.num))
                    num.AddAnnotation(new FromPreviousSourceSemaphore());
            }

            if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null &&
                sourceDocument.MainDocumentPart.NumberingDefinitionsPart != null)
            {
                AddRelationships(sourceDocument.MainDocumentPart.NumberingDefinitionsPart,
                    newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart.NumberingDefinitionsPart,
                    newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }
        }

        // Note: the following two methods were added with almost exact duplicate code to the method above, because I do not want to touch that code.
        private static void CopyNumberingForGlossaryDocumentPartFromGD(NumberingDefinitionsPart sourceNumberingPart, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var numIdMap = new Dictionary<int, int>();
            var number = 1;
            var abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (var numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                var idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    oldNumbering ??= sourceNumberingPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            var numIds = newNumbering
                                .Root
                                .Elements(W.num)
                                .Select(f => (int)f.Attribute(W.numId));
                            if (numIds.Any())
                                number = numIds.Max() + 1;
                            numIds = newNumbering
                                .Root
                                .Elements(W.abstractNum)
                                .Select(f => (int)f.Attribute(W.abstractNumId));
                            if (numIds.Any())
                                abstractNumber = numIds.Max() + 1;
                        }
                        else
                        {
                            newDocument.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            newNumbering.Declaration.Standalone = Yes;
                            newNumbering.Declaration.Encoding = Utf8;
                            newNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                        }
                    }
                    var numId = (int)idElement.Attribute(W.val);
                    if (numId != 0)
                    {
                        var element = oldNumbering
                            .Descendants(W.num)
                            .Where(p => ((int)p.Attribute(W.numId)) == numId)
                            .FirstOrDefault();
                        if (element == null)
                            continue;

                        // Copy abstract numbering element, if necessary (use matching NSID)
                        var abstractNumIdStr = (string)element
                            .Elements(W.abstractNumId)
                            .First()
                            .Attribute(W.val);
                        if (!int.TryParse(abstractNumIdStr, out var abstractNumId))
                            throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");
                        var abstractElement = oldNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                            .First();
                        var nsidElement = abstractElement
                            .Element(W.nsid);
                        string abstractNSID = null;
                        if (nsidElement != null)
                            abstractNSID = (string)nsidElement
                                .Attribute(W.val);
                        var newAbstractElement = newNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                            .Where(p =>
                            {
                                var thisNsidElement = p.Element(W.nsid);
                                if (thisNsidElement == null)
                                    return false;
                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                            })
                            .FirstOrDefault();
                        if (newAbstractElement == null)
                        {
                            newAbstractElement = new XElement(abstractElement);
                            newAbstractElement.SetAttributeValue(W.abstractNumId, abstractNumber.ToString());
                            abstractNumber++;
                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                            else
                                newNumbering.Root.Add(newAbstractElement);

                            foreach (var pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                            {
                                var bulletId = (string)pictId.Attribute(W.val);
                                var numPicBullet = oldNumbering
                                    .Descendants(W.numPicBullet)
                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                var maxNumPicBulletId = new int[] { -1 }.Concat(
                                    newNumbering.Descendants(W.numPicBullet)
                                    .Attributes(W.numPicBulletId)
                                    .Select(a => (int)a))
                                    .Max() + 1;
                                var newNumPicBullet = new XElement(numPicBullet);
                                newNumPicBullet.SetAttributeValue(W.numPicBulletId, maxNumPicBulletId.ToString());
                                pictId.SetAttributeValue(W.val, maxNumPicBulletId.ToString());
                                newNumbering.Root.AddFirst(newNumPicBullet);
                            }
                        }
                        var newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                        // Copy numbering element, if necessary (use matching element with no overrides)
                        XElement newElement;
                        if (numIdMap.ContainsKey(numId))
                        {
                            newElement = newNumbering
                                .Descendants()
                                .Elements(W.num)
                                .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                .First();
                        }
                        else
                        {
                            newElement = new XElement(element);
                            newElement
                                .Elements(W.abstractNumId)
                                .First()
                                .SetAttributeValue(W.val, newAbstractId);
                            newElement.SetAttributeValue(W.numId, number.ToString());
                            numIdMap.Add(numId, number);
                            number++;
                            newNumbering.Root.Add(newElement);
                        }
                        idElement.SetAttributeValue(W.val, newElement.Attribute(W.numId).Value);
                    }
                }
            }
            if (newNumbering != null)
            {
                foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                    abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                foreach (var num in newNumbering.Descendants(W.num))
                    num.AddAnnotation(new FromPreviousSourceSemaphore());
            }

            if (newDocument.MainDocumentPart.NumberingDefinitionsPart != null &&
                sourceNumberingPart != null)
            {
                AddRelationships(sourceNumberingPart,
                    newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceNumberingPart,
                    newDocument.MainDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }

            newDocument.MainDocumentPart.NumberingDefinitionsPart?.PutXDocument();
        }

        private static void CopyNumberingForGlossaryDocumentPartToGD(NumberingDefinitionsPart sourceNumberingPart, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var numIdMap = new Dictionary<int, int>();
            var number = 1;
            var abstractNumber = 0;
            XDocument oldNumbering = null;
            XDocument newNumbering = null;

            foreach (var numReference in newContent.DescendantsAndSelf(W.numPr))
            {
                var idElement = numReference.Descendants(W.numId).FirstOrDefault();
                if (idElement != null)
                {
                    oldNumbering ??= sourceNumberingPart.GetXDocument();
                    if (newNumbering == null)
                    {
                        if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null)
                        {
                            newNumbering = newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            var numIds = newNumbering
                                .Root
                                .Elements(W.num)
                                .Select(f => (int)f.Attribute(W.numId));
                            if (numIds.Any())
                                number = numIds.Max() + 1;
                            numIds = newNumbering
                                .Root
                                .Elements(W.abstractNum)
                                .Select(f => (int)f.Attribute(W.abstractNumId));
                            if (numIds.Any())
                                abstractNumber = numIds.Max() + 1;
                        }
                        else
                        {
                            newDocument.MainDocumentPart.GlossaryDocumentPart.AddNewPart<NumberingDefinitionsPart>();
                            newNumbering = newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument();
                            newNumbering.Declaration.Standalone = Yes;
                            newNumbering.Declaration.Encoding = Utf8;
                            newNumbering.Add(new XElement(W.numbering, NamespaceAttributes));
                        }
                    }
                    var numId = (int)idElement.Attribute(W.val);
                    if (numId != 0)
                    {
                        var element = oldNumbering
                            .Descendants(W.num)
                            .Where(p => ((int)p.Attribute(W.numId)) == numId)
                            .FirstOrDefault();
                        if (element == null)
                            continue;

                        // Copy abstract numbering element, if necessary (use matching NSID)
                        var abstractNumIdStr = (string)element
                            .Elements(W.abstractNumId)
                            .First()
                            .Attribute(W.val);
                        if (!int.TryParse(abstractNumIdStr, out var abstractNumId))
                            throw new DocumentBuilderException("Invalid document - invalid value for abstractNumId");
                        var abstractElement = oldNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(p => ((int)p.Attribute(W.abstractNumId)) == abstractNumId)
                            .First();
                        var nsidElement = abstractElement
                            .Element(W.nsid);
                        string abstractNSID = null;
                        if (nsidElement != null)
                            abstractNSID = (string)nsidElement
                                .Attribute(W.val);
                        var newAbstractElement = newNumbering
                            .Descendants()
                            .Elements(W.abstractNum)
                            .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                            .Where(p =>
                            {
                                var thisNsidElement = p.Element(W.nsid);
                                if (thisNsidElement == null)
                                    return false;
                                return (string)thisNsidElement.Attribute(W.val) == abstractNSID;
                            })
                            .FirstOrDefault();
                        if (newAbstractElement == null)
                        {
                            newAbstractElement = new XElement(abstractElement);
                            newAbstractElement.SetAttributeValue(W.abstractNumId, abstractNumber.ToString());
                            abstractNumber++;
                            if (newNumbering.Root.Elements(W.abstractNum).Any())
                                newNumbering.Root.Elements(W.abstractNum).Last().AddAfterSelf(newAbstractElement);
                            else
                                newNumbering.Root.Add(newAbstractElement);

                            foreach (var pictId in newAbstractElement.Descendants(W.lvlPicBulletId))
                            {
                                var bulletId = (string)pictId.Attribute(W.val);
                                var numPicBullet = oldNumbering
                                    .Descendants(W.numPicBullet)
                                    .FirstOrDefault(d => (string)d.Attribute(W.numPicBulletId) == bulletId);
                                var maxNumPicBulletId = new int[] { -1 }.Concat(
                                    newNumbering.Descendants(W.numPicBullet)
                                    .Attributes(W.numPicBulletId)
                                    .Select(a => (int)a))
                                    .Max() + 1;
                                var newNumPicBullet = new XElement(numPicBullet);
                                newNumPicBullet.SetAttributeValue(W.numPicBulletId, maxNumPicBulletId.ToString());
                                pictId.SetAttributeValue(W.val, maxNumPicBulletId.ToString());
                                newNumbering.Root.AddFirst(newNumPicBullet);
                            }
                        }
                        var newAbstractId = newAbstractElement.Attribute(W.abstractNumId).Value;

                        // Copy numbering element, if necessary (use matching element with no overrides)
                        XElement newElement;
                        if (numIdMap.ContainsKey(numId))
                        {
                            newElement = newNumbering
                                .Descendants()
                                .Elements(W.num)
                                .Where(e => e.Annotation<FromPreviousSourceSemaphore>() == null)
                                .Where(p => ((int)p.Attribute(W.numId)) == numIdMap[numId])
                                .First();
                        }
                        else
                        {
                            newElement = new XElement(element);
                            newElement
                                .Elements(W.abstractNumId)
                                .First()
                                .SetAttributeValue(W.val, newAbstractId);
                            newElement.SetAttributeValue(W.numId, number.ToString());
                            numIdMap.Add(numId, number);
                            number++;
                            newNumbering.Root.Add(newElement);
                        }
                        idElement.SetAttributeValue(W.val, newElement.Attribute(W.numId).Value);
                    }
                }
            }
            if (newNumbering != null)
            {
                foreach (var abstractNum in newNumbering.Descendants(W.abstractNum))
                    abstractNum.AddAnnotation(new FromPreviousSourceSemaphore());
                foreach (var num in newNumbering.Descendants(W.num))
                    num.AddAnnotation(new FromPreviousSourceSemaphore());
            }

            if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null &&
                sourceNumberingPart != null)
            {
                AddRelationships(sourceNumberingPart,
                    newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceNumberingPart,
                    newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart,
                    new[] { newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.GetXDocument().Root }, images);
            }
            if (newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart != null)
                newDocument.MainDocumentPart.GlossaryDocumentPart.NumberingDefinitionsPart.PutXDocument();
        }

        private static void CopyRelatedImage(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName,
            List<ImageData> images)
        {
            var relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;

            // First look to see if this relId has already been added to the new document.
            // This is necessary for those parts that get processed with both old and new ids, such as the comments
            // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
            // in that case.
            var tempPartIdPair5 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (tempPartIdPair5 != null)
                return;

            var tempEr5 = newContentPart.ExternalRelationships.FirstOrDefault(er => er.Id == relId);
            if (tempEr5 != null)
                return;

            var ipp2 = oldContentPart.Parts.FirstOrDefault(ipp => ipp.RelationshipId == relId);
            if (ipp2 != null)
            {
                var oldPart2 = ipp2.OpenXmlPart;
                if (!(oldPart2 is ImagePart))
                    throw new DocumentBuilderException("Invalid document - target part is not ImagePart");

                var oldPart = (ImagePart)ipp2.OpenXmlPart;
                var temp = ManageImageCopy(oldPart, newContentPart, images);
                if (temp.ImagePart == null)
                {
                    var newPart = newContentPart switch
                    {
                        MainDocumentPart part => part.AddImagePart(oldPart.ContentType),
                        HeaderPart part => part.AddImagePart(oldPart.ContentType),
                        FooterPart part => part.AddImagePart(oldPart.ContentType),
                        EndnotesPart part => part.AddImagePart(oldPart.ContentType),
                        FootnotesPart part => part.AddImagePart(oldPart.ContentType),
                        ThemePart part => part.AddImagePart(oldPart.ContentType),
                        WordprocessingCommentsPart part => part.AddImagePart(oldPart.ContentType),
                        DocumentSettingsPart part => part.AddImagePart(oldPart.ContentType),
                        ChartPart part => part.AddImagePart(oldPart.ContentType),
                        NumberingDefinitionsPart part => part.AddImagePart(oldPart.ContentType),
                        DiagramDataPart part => part.AddImagePart(oldPart.ContentType),
                        ChartDrawingPart part => part.AddImagePart(oldPart.ContentType),
                        _ => null
                    };
                    temp.ImagePart = newPart;
                    var id = newContentPart.GetIdOfPart(newPart);
                    temp.AddContentPartRelTypeResourceIdTupple(newContentPart, newPart.RelationshipType, id);
                    imageReference.SetAttributeValue(attributeName, id);
                    temp.WriteImage(newPart);
                }
                else
                {
                    var refRel = newContentPart.Parts.FirstOrDefault(pip =>
                    {
                        var rel = temp.ContentPartRelTypeIdList.FirstOrDefault(cpr =>
                        {
                            var found = cpr.ContentPart == newContentPart;
                            return found;
                        });
                        return rel != null;
                    });
                    if (refRel != null)
                    {
                        imageReference.SetAttributeValue(attributeName, temp.ContentPartRelTypeIdList.First(cpr =>
                        {
                            var found = cpr.ContentPart == newContentPart;
                            return found;
                        }).RelationshipId);
                        return;
                    }
                    var newId = Relationships.GetNewRelationshipId();
                    newContentPart.CreateRelationshipToPart(temp.ImagePart, newId);
                    imageReference.SetAttributeValue(R.id, newId);
                }
            }
            else
            {
                var er = oldContentPart.ExternalRelationships.FirstOrDefault(er1 => er1.Id == relId);
                if (er != null)
                {
                    var newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    imageReference.SetAttributeValue(R.id, newEr.Id);
                    return;
                }
                throw new DocumentBuilderInternalException("Source {0} is unsupported document - contains reference to NULL image");
            }
        }

        private static void CopyRelatedPartsForContentParts(OpenXmlPart oldContentPart, OpenXmlPart newContentPart,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => d.Name == VML.imagedata || d.Name == VML.fill || d.Name == VML.stroke || d.Name == A.blip)
                .ToList();
            foreach (var imageReference in relevantElements)
            {
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.embed, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.pict, images);
                CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.id, images);
            }

            foreach (var diagramReference in newContent.DescendantsAndSelf().Where(d => d.Name == DGM.relIds || d.Name == A.relIds))
            {
                // dm attribute
                var relId = diagramReference.Attribute(R.dm).Value;
                var ipp = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp != null)
                {
                    var tempPart = ipp.OpenXmlPart;
                    continue;
                }

                var tempEr = newContentPart.ExternalRelationships.FirstOrDefault(er2 => er2.Id == relId);
                if (tempEr != null)
                    continue;

                var oldPart = oldContentPart.GetPartById(relId);
                OpenXmlPart newPart = newContentPart.AddNewPart<DiagramDataPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.dm).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // lo attribute
                relId = diagramReference.Attribute(R.lo).Value;
                var ipp2 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp2 != null)
                {
                    var tempPart = ipp2.OpenXmlPart;
                    continue;
                }


                var tempEr4 = newContentPart.ExternalRelationships.FirstOrDefault(er3 => er3.Id == relId);
                if (tempEr4 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramLayoutDefinitionPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.lo).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // qs attribute
                relId = diagramReference.Attribute(R.qs).Value;
                var ipp5 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp5 != null)
                {
                    var tempPart = ipp5.OpenXmlPart;
                    continue;
                }

                var tempEr5 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr5 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramStylePart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.qs).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);

                // cs attribute
                relId = diagramReference.Attribute(R.cs).Value;
                var ipp6 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp6 != null)
                {
                    var tempPart = ipp6.OpenXmlPart;
                    continue;
                }

                var tempEr6 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr6 != null)
                    continue;

                oldPart = oldContentPart.GetPartById(relId);
                newPart = newContentPart.AddNewPart<DiagramColorsPart>();
                newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
                diagramReference.Attribute(R.cs).Value = newContentPart.GetIdOfPart(newPart);
                AddRelationships(oldPart, newPart, new[] { newPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newPart.GetXDocument().Root }, images);
            }

            foreach (var oleReference in newContent.DescendantsAndSelf(O.OLEObject))
            {
                var relId = (string)oleReference.Attribute(R.id);

                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                var ipp1 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp1 != null)
                {
                    var tempPart = ipp1.OpenXmlPart;
                    continue;
                }

                var tempEr1 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr1 != null)
                    continue;

                var ipp4 = oldContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp4 != null)
                {
                    var oldPart = oldContentPart.GetPartById(relId);
                    OpenXmlPart newPart = null;
                    newPart = oldPart switch
                    {
                        EmbeddedObjectPart _ => newContentPart switch
                        {
                            HeaderPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            FooterPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            MainDocumentPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            FootnotesPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            EndnotesPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            WordprocessingCommentsPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                            _ => newPart
                        },
                        EmbeddedPackagePart _ => newContentPart switch
                        {
                            HeaderPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            FooterPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            MainDocumentPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            FootnotesPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            EndnotesPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            WordprocessingCommentsPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            ChartPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                            _ => newPart
                        },
                        _ => newPart
                    };
                    using (var oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                    using (var newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                    {
                        oldObject.CopyTo(newObject);
                    }
                    oleReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                }
                else
                {
                    if (relId != null)
                    {
                        var er = oldContentPart.GetExternalRelationship(relId);
                        var newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                        oleReference.SetAttributeValue(R.id, newEr.Id);
                    }
                }
            }

            foreach (var chartReference in newContent.DescendantsAndSelf(C.chart))
            {
                var relId = (string)chartReference.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;
                var ipp2 = newContentPart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp2 != null)
                {
                    var tempPart = ipp2.OpenXmlPart;
                    continue;
                }

                var tempEr2 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr2 != null)
                    continue;

                var ipp3 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp3 == null)
                    continue;
                var oldPart = (ChartPart)ipp3.OpenXmlPart;
                var oldChart = oldPart.GetXDocument();
                var newPart = newContentPart.AddNewPart<ChartPart>();
                var newChart = newPart.GetXDocument();
                newChart.Add(oldChart.Root);
                chartReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                CopyChartObjects(oldPart, newPart);
                CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newChart.Root }, images);
            }

            foreach (var userShape in newContent.DescendantsAndSelf(C.userShapes))
            {
                var relId = (string)userShape.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var ipp4 = newContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp4 != null)
                {
                    var tempPart = ipp4.OpenXmlPart;
                    continue;
                }

                var tempEr4 = newContentPart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr4 != null)
                    continue;

                var ipp5 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
                if (ipp5 != null)
                {
                    var oldPart = (ChartDrawingPart)ipp5.OpenXmlPart;
                    var oldXDoc = oldPart.GetXDocument();
                    var newPart = newContentPart.AddNewPart<ChartDrawingPart>();
                    var newXDoc = newPart.GetXDocument();
                    newXDoc.Add(oldXDoc.Root);
                    userShape.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                    AddRelationships(oldPart, newPart, newContent);
                    CopyRelatedPartsForContentParts(oldPart, newPart, new[] { newXDoc.Root }, images);
                }
            }
        }

        private static void CopyFontTable(FontTablePart oldFontTablePart, FontTablePart newFontTablePart)
        {
            var relevantElements = oldFontTablePart.GetXDocument().Descendants().Where(d => d.Name == W.embedRegular ||
                d.Name == W.embedBold || d.Name == W.embedItalic || d.Name == W.embedBoldItalic).ToList();
            foreach (var fontReference in relevantElements)
            {
                var relId = (string)fontReference.Attribute(R.id);
                if (string.IsNullOrEmpty(relId))
                    continue;

                var ipp1 = newFontTablePart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp1 != null)
                {
                    var tempPart = ipp1.OpenXmlPart;
                    continue;
                }

                var tempEr1 = newFontTablePart.ExternalRelationships.FirstOrDefault(z => z.Id == relId);
                if (tempEr1 != null)
                    continue;

                var oldPart2 = oldFontTablePart.GetPartById(relId);
                if (oldPart2 == null || (!(oldPart2 is FontPart)))
                    throw new DocumentBuilderException("Invalid document - FontTablePart contains invalid relationship");

                var oldPart = (FontPart)oldPart2;
                var newPart = newFontTablePart.AddFontPart(oldPart.ContentType);
                var resourceId = newFontTablePart.GetIdOfPart(newPart);
                using (var oldFont = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                using (var newFont = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                {
                    oldFont.CopyTo(newFont);
                }
                fontReference.SetAttributeValue(R.id, resourceId);
            }
        }

        private static void CopyChartObjects(ChartPart oldChart, ChartPart newChart)
        {
            foreach (var dataReference in newChart.GetXDocument().Descendants(C.externalData))
            {
                var relId = dataReference.Attribute(R.id).Value;

                var ipp1 = oldChart.Parts.FirstOrDefault(z => z.RelationshipId == relId);
                if (ipp1 != null)
                {
                    var oldRelatedPart = ipp1.OpenXmlPart;
                    if (oldRelatedPart is EmbeddedPackagePart)
                    {
                        var oldPart = (EmbeddedPackagePart)ipp1.OpenXmlPart;
                        var newPart = newChart.AddEmbeddedPackagePart(oldPart.ContentType);
                        using (var oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                        using (var newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                        {
                            oldObject.CopyTo(newObject);
                        }
                        dataReference.SetAttributeValue(R.id, newChart.GetIdOfPart(newPart));
                    }
                    else if (oldRelatedPart is EmbeddedObjectPart)
                    {
                        var oldPart = (EmbeddedObjectPart)ipp1.OpenXmlPart;
                        var relType = oldRelatedPart.RelationshipType;
                        var conType = oldRelatedPart.ContentType;
                        var id = Relationships.GetNewRelationshipId();
                        var newPart = newChart.AddExtendedPart(relType, conType, ".bin", id);
                        using (var oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                        using (var newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite))
                        {
                            oldObject.CopyTo(newObject);
                        }
                        dataReference.SetAttributeValue(R.id, newChart.GetIdOfPart(newPart));
                    }
                }
                else
                {
                    var oldRelationship = oldChart.GetExternalRelationship(relId);
                    var newRid = Relationships.GetNewRelationshipId();
                    var oldRel =
                        oldChart.ExternalRelationships.FirstOrDefault(h => h.Id == relId)
                        ?? throw new DocumentBuilderInternalException("Internal Error 0007");
                    newChart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    dataReference.SetAttributeValue(R.id, newRid);
                }
            }
        }

        private static void CopyStartingParts(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            List<ImageData> images)
        {
            // A Core File Properties part does not have implicit or explicit relationships to other parts.
            var corePart = sourceDocument.CoreFilePropertiesPart;
            if (corePart?.GetXDocument().Root != null)
            {
                newDocument.AddCoreFilePropertiesPart();
                var newXDoc = newDocument.CoreFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                var sourceXDoc = corePart.GetXDocument();
                newXDoc.Add(sourceXDoc.Root);
            }

            // An application attributes part does not have implicit or explicit relationships to other parts.
            var extPart = sourceDocument.ExtendedFilePropertiesPart;
            if (extPart != null)
            {
                OpenXmlPart newPart = newDocument.AddExtendedFilePropertiesPart();
                var newXDoc = newDocument.ExtendedFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(extPart.GetXDocument().Root);
            }

            // An custom file properties part does not have implicit or explicit relationships to other parts.
            var customPart = sourceDocument.CustomFilePropertiesPart;
            if (customPart != null)
            {
                newDocument.AddCustomFilePropertiesPart();
                var newXDoc = newDocument.CustomFilePropertiesPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;

                // Remove custom properties (source doc metadata irrelevant for generated document)
                var propsDocument = customPart.GetXDocument().Root;
                if (propsDocument?.HasElements == true)
                    propsDocument.RemoveNodes();
                newXDoc.Add(propsDocument);
            }

            var oldSettingsPart = sourceDocument.MainDocumentPart.DocumentSettingsPart;
            if (oldSettingsPart != null)
            {
                var newSettingsPart = newDocument.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                var settingsXDoc = oldSettingsPart.GetXDocument();
                AddRelationships(oldSettingsPart, newSettingsPart, new[] { settingsXDoc.Root });
                CopyFootnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                CopyEndnotesPart(sourceDocument, newDocument, settingsXDoc, images);
                var newXDoc = newDocument.MainDocumentPart.DocumentSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
                CopyRelatedPartsForContentParts(oldSettingsPart, newSettingsPart, new[] { newXDoc.Root }, images);
            }

            var oldWebSettingsPart = sourceDocument.MainDocumentPart.WebSettingsPart;
            if (oldWebSettingsPart != null)
            {
                var newWebSettingsPart = newDocument.MainDocumentPart.AddNewPart<WebSettingsPart>();
                var settingsXDoc = oldWebSettingsPart.GetXDocument();
                AddRelationships(oldWebSettingsPart, newWebSettingsPart, new[] { settingsXDoc.Root });
                var newXDoc = newDocument.MainDocumentPart.WebSettingsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(settingsXDoc.Root);
            }

            var themePart = sourceDocument.MainDocumentPart.ThemePart;
            if (themePart != null)
            {
                var newThemePart = newDocument.MainDocumentPart.AddNewPart<ThemePart>();
                var newXDoc = newDocument.MainDocumentPart.ThemePart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(themePart.GetXDocument().Root);
                CopyRelatedPartsForContentParts(themePart, newThemePart, new[] { newThemePart.GetXDocument().Root }, images);
            }

            // If needed to handle GlossaryDocumentPart in the future, then
            // this code should handle the following parts:
            //   MainDocumentPart.GlossaryDocumentPart.StyleDefinitionsPart
            //   MainDocumentPart.GlossaryDocumentPart.StylesWithEffectsPart

            // A Style Definitions part shall not have implicit or explicit relationships to any other part.
            var stylesPart = sourceDocument.MainDocumentPart.StyleDefinitionsPart;
            if (stylesPart != null)
            {
                newDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();
                var newXDoc = newDocument.MainDocumentPart.StyleDefinitionsPart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                newXDoc.Add(new XElement(W.styles,
                    new XAttribute(XNamespace.Xmlns + "w", W.w)

                    //,
                    //stylesPart.GetXDocument().Descendants(W.docDefaults)

                    //,
                    //new XElement(W.latentStyles, stylesPart.GetXDocument().Descendants(W.latentStyles).Attributes())

                    ));
                MergeDocDefaultStyles(stylesPart.GetXDocument(), newXDoc);
                MergeStyles(sourceDocument, newDocument, stylesPart.GetXDocument(), newXDoc, Enumerable.Empty<XElement>());
                MergeLatentStyles(stylesPart.GetXDocument(), newXDoc);
            }

            // A Font Table part shall not have any implicit or explicit relationships to any other part.
            var fontTablePart = sourceDocument.MainDocumentPart.FontTablePart;
            if (fontTablePart != null)
            {
                newDocument.MainDocumentPart.AddNewPart<FontTablePart>();
                var newXDoc = newDocument.MainDocumentPart.FontTablePart.GetXDocument();
                newXDoc.Declaration.Standalone = Yes;
                newXDoc.Declaration.Encoding = Utf8;
                CopyFontTable(sourceDocument.MainDocumentPart.FontTablePart, newDocument.MainDocumentPart.FontTablePart);
                newXDoc.Add(fontTablePart.GetXDocument().Root);
            }
        }

        private static void CopyFootnotesPart(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            XDocument settingsXDoc, List<ImageData> images)
        {
            var number = 0;
            XDocument oldFootnotes = null;
            XDocument newFootnotes = null;
            var footnotePr = settingsXDoc.Root.Element(W.footnotePr);
            if (footnotePr == null)
                return;
            if (sourceDocument.MainDocumentPart.FootnotesPart == null)
                return;
            foreach (var footnote in footnotePr.Elements(W.footnote))
            {
                oldFootnotes ??= sourceDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                if (newFootnotes == null)
                {
                    if (newDocument.MainDocumentPart.FootnotesPart != null)
                    {
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.Standalone = Yes;
                        newFootnotes.Declaration.Encoding = Utf8;
                        var ids = newFootnotes.Root.Elements(W.footnote).Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<FootnotesPart>();
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.Standalone = Yes;
                        newFootnotes.Declaration.Encoding = Utf8;
                        newFootnotes.Add(new XElement(W.footnotes, NamespaceAttributes));
                    }
                }
                var id = (string)footnote.Attribute(W.id);
                var element = oldFootnotes.Descendants()
                    .Elements(W.footnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    var newElement = new XElement(element);
                    // the following adds the footnote into the new settings part
                    newElement.SetAttributeValue(W.id, number.ToString());
                    newFootnotes.Root.Add(newElement);
                    footnote.SetAttributeValue(W.id, number.ToString());
                    number++;
                }
            }
        }

        private static void CopyEndnotesPart(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            XDocument settingsXDoc, List<ImageData> images)
        {
            var number = 0;
            XDocument oldEndnotes = null;
            XDocument newEndnotes = null;
            var endnotePr = settingsXDoc.Root.Element(W.endnotePr);
            if (endnotePr == null)
                return;
            if (sourceDocument.MainDocumentPart.EndnotesPart == null)
                return;
            foreach (var endnote in endnotePr.Elements(W.endnote))
            {
                oldEndnotes ??= sourceDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                if (newEndnotes == null)
                {
                    if (newDocument.MainDocumentPart.EndnotesPart != null)
                    {
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.Standalone = Yes;
                        newEndnotes.Declaration.Encoding = Utf8;
                        var ids = newEndnotes.Root
                            .Elements(W.endnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<EndnotesPart>();
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.Standalone = Yes;
                        newEndnotes.Declaration.Encoding = Utf8;
                        newEndnotes.Add(new XElement(W.endnotes, NamespaceAttributes));
                    }
                }
                var id = (string)endnote.Attribute(W.id);
                var element = oldEndnotes.Descendants()
                    .Elements(W.endnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    var newElement = new XElement(element);
                    newElement.SetAttributeValue(W.id, number.ToString());
                    newEndnotes.Root.Add(newElement);
                    endnote.SetAttributeValue(W.id, number.ToString());
                    number++;
                }
            }
        }

        public static void FixRanges(XDocument sourceDocument, IEnumerable<XElement> newContent)
        {
            FixRange(sourceDocument,
                newContent,
                W.commentRangeStart,
                W.commentRangeEnd,
                W.id,
                W.commentReference);
            FixRange(sourceDocument,
                newContent,
                W.bookmarkStart,
                W.bookmarkEnd,
                W.id,
                null);
            FixRange(sourceDocument,
                newContent,
                W.permStart,
                W.permEnd,
                W.id,
                null);
            FixRange(sourceDocument,
                newContent,
                W.moveFromRangeStart,
                W.moveFromRangeEnd,
                W.id,
                null);
            FixRange(sourceDocument,
                newContent,
                W.moveToRangeStart,
                W.moveToRangeEnd,
                W.id,
                null);
            DeleteUnmatchedRange(sourceDocument,
                newContent,
                W.moveFromRangeStart,
                W.moveFromRangeEnd,
                W.moveToRangeStart,
                W.name,
                W.id);
            DeleteUnmatchedRange(sourceDocument,
                newContent,
                W.moveToRangeStart,
                W.moveToRangeEnd,
                W.moveFromRangeStart,
                W.name,
                W.id);
        }

        private static void AddAtBeginning(IEnumerable<XElement> newContent, XElement contentToAdd)
        {
            var first = newContent.First();
            if (first.Element(W.pPr) != null)
                first.Element(W.pPr).AddAfterSelf(contentToAdd);
            else
                first.AddFirst(new XElement(contentToAdd));
        }

        private static void AddAtEnd(IEnumerable<XElement> newContent, XElement contentToAdd)
        {
            var last = newContent.Last();
            if (last.Element(W.pPr) != null)
                last.Element(W.pPr).AddAfterSelf(new XElement(contentToAdd));
            else
                last.Add(new XElement(contentToAdd));
        }

        // If the set of paragraphs from sourceDocument don't have a complete start/end for bookmarks,
        // comments, etc., then this adds them to the paragraph.  Note that this adds them to
        // sourceDocument, and is impure.
        private static void FixRange(XDocument sourceDocument, IEnumerable<XElement> newContent,
            XName startElement, XName endElement, XName idAttribute, XName refElement)
        {
            foreach (var start in newContent.DescendantsAndSelf(startElement))
            {
                var rangeId = start.Attribute(idAttribute).Value;
                if (newContent
                    .DescendantsAndSelf(endElement)
                    .Where(e => e.Attribute(idAttribute).Value == rangeId)
                    .Count() == 0)
                {
                    var end = sourceDocument
                        .Descendants(endElement)
                        .Where(o => o.Attribute(idAttribute).Value == rangeId)
                        .FirstOrDefault();
                    if (end != null)
                    {
                        AddAtEnd(newContent, new XElement(end));
                        if (refElement != null)
                        {
                            var newRef = new XElement(refElement, new XAttribute(idAttribute, rangeId));
                            AddAtEnd(newContent, new XElement(newRef));
                        }
                    }
                }
            }
            foreach (var end in newContent.Elements(endElement))
            {
                var rangeId = end.Attribute(idAttribute).Value;
                if (newContent
                    .DescendantsAndSelf(startElement)
                    .Where(s => s.Attribute(idAttribute).Value == rangeId)
                    .Count() == 0)
                {
                    var start = sourceDocument
                        .Descendants(startElement)
                        .Where(o => o.Attribute(idAttribute).Value == rangeId)
                        .FirstOrDefault();
                    if (start != null)
                        AddAtBeginning(newContent, new XElement(start));
                }
            }
        }

        private static void DeleteUnmatchedRange(XDocument sourceDocument, IEnumerable<XElement> newContent,
            XName startElement, XName endElement, XName matchTo, XName matchAttr, XName idAttr)
        {
            var deleteList = new List<string>();
            foreach (var start in newContent.Elements(startElement))
            {
                var id = start.Attribute(matchAttr).Value;
                if (!newContent.Elements(matchTo).Where(n => n.Attribute(matchAttr).Value == id).Any())
                    deleteList.Add(start.Attribute(idAttr).Value);
            }
            foreach (var item in deleteList)
            {
                newContent.Elements(startElement).Where(n => n.Attribute(idAttr).Value == item).Remove();
                newContent.Elements(endElement).Where(n => n.Attribute(idAttr).Value == item).Remove();
                newContent.Where(p => p.Name == startElement && p.Attribute(idAttr).Value == item).Remove();
                newContent.Where(p => p.Name == endElement && p.Attribute(idAttr).Value == item).Remove();
            }
        }

        private static void CopyFootnotes(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var number = 0;
            XDocument oldFootnotes = null;
            XDocument newFootnotes = null;
            foreach (var footnote in newContent.DescendantsAndSelf(W.footnoteReference))
            {
                oldFootnotes ??= sourceDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                if (newFootnotes == null)
                {
                    if (newDocument.MainDocumentPart.FootnotesPart != null)
                    {
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        var ids = newFootnotes
                            .Root
                            .Elements(W.footnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<FootnotesPart>();
                        newFootnotes = newDocument.MainDocumentPart.FootnotesPart.GetXDocument();
                        newFootnotes.Declaration.Standalone = Yes;
                        newFootnotes.Declaration.Encoding = Utf8;
                        newFootnotes.Add(new XElement(W.footnotes, NamespaceAttributes));
                    }
                }
                var id = (string)footnote.Attribute(W.id);
                var element = oldFootnotes
                    .Descendants()
                    .Elements(W.footnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .FirstOrDefault();
                if (element != null)
                {
                    var newElement = new XElement(element);
                    newElement.SetAttributeValue(W.id, number.ToString());
                    newFootnotes.Root.Add(newElement);
                    footnote.SetAttributeValue(W.id, number.ToString());
                    number++;
                }
            }
            if (sourceDocument.MainDocumentPart.FootnotesPart != null &&
                newDocument.MainDocumentPart.FootnotesPart != null)
            {
                AddRelationships(sourceDocument.MainDocumentPart.FootnotesPart,
                    newDocument.MainDocumentPart.FootnotesPart,
                    new[] { newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart.FootnotesPart,
                    newDocument.MainDocumentPart.FootnotesPart,
                    new[] { newDocument.MainDocumentPart.FootnotesPart.GetXDocument().Root }, images);
            }
        }

        private static void CopyEndnotes(WordprocessingDocument sourceDocument, WordprocessingDocument newDocument,
            IEnumerable<XElement> newContent, List<ImageData> images)
        {
            var number = 0;
            XDocument oldEndnotes = null;
            XDocument newEndnotes = null;
            foreach (var endnote in newContent.DescendantsAndSelf(W.endnoteReference))
            {
                oldEndnotes ??= sourceDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                if (newEndnotes == null)
                {
                    if (newDocument.MainDocumentPart.EndnotesPart != null)
                    {
                        newEndnotes = newDocument
                            .MainDocumentPart
                            .EndnotesPart
                            .GetXDocument();
                        var ids = newEndnotes
                            .Root
                            .Elements(W.endnote)
                            .Select(f => (int)f.Attribute(W.id));
                        if (ids.Any())
                            number = ids.Max() + 1;
                    }
                    else
                    {
                        newDocument.MainDocumentPart.AddNewPart<EndnotesPart>();
                        newEndnotes = newDocument.MainDocumentPart.EndnotesPart.GetXDocument();
                        newEndnotes.Declaration.Standalone = Yes;
                        newEndnotes.Declaration.Encoding = Utf8;
                        newEndnotes.Add(new XElement(W.endnotes, NamespaceAttributes));
                    }
                }
                var id = (string)endnote.Attribute(W.id);
                var element = oldEndnotes
                    .Descendants()
                    .Elements(W.endnote)
                    .Where(p => ((string)p.Attribute(W.id)) == id)
                    .First();
                var newElement = new XElement(element);
                newElement.SetAttributeValue(W.id, number.ToString());
                newEndnotes.Root.Add(newElement);
                endnote.SetAttributeValue(W.id, number.ToString());
                number++;
            }
            if (sourceDocument.MainDocumentPart.EndnotesPart != null &&
                newDocument.MainDocumentPart.EndnotesPart != null)
            {
                AddRelationships(sourceDocument.MainDocumentPart.EndnotesPart,
                    newDocument.MainDocumentPart.EndnotesPart,
                    new[] { newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Root });
                CopyRelatedPartsForContentParts(sourceDocument.MainDocumentPart.EndnotesPart,
                    newDocument.MainDocumentPart.EndnotesPart,
                    new[] { newDocument.MainDocumentPart.EndnotesPart.GetXDocument().Root }, images);
            }
        }

        // General function for handling images that tries to use an existing image if they are the same
        private static ImageData ManageImageCopy(ImagePart oldImage, OpenXmlPart newContentPart, List<ImageData> images)
        {
            var oldImageData = new ImageData(oldImage);
            foreach (var item in images)
            {
                if (newContentPart != item.ImagePart)
                    continue;
                if (item.Compare(oldImageData))
                    return item;
            }
            images.Add(oldImageData);
            return oldImageData;
        }

        private static readonly XAttribute[] NamespaceAttributes =
        {
            new(XNamespace.Xmlns + "wpc", WPC.wpc),
            new(XNamespace.Xmlns + "mc", MC.mc),
            new(XNamespace.Xmlns + "o", O.o),
            new(XNamespace.Xmlns + "r", R.r),
            new(XNamespace.Xmlns + "m", M.m),
            new(XNamespace.Xmlns + "v", VML.vml),
            new(XNamespace.Xmlns + "wp14", WP14.wp14),
            new(XNamespace.Xmlns + "wp", WP.wp),
            new(XNamespace.Xmlns + "w10", W10.w10),
            new(XNamespace.Xmlns + "w", W.w),
            new(XNamespace.Xmlns + "w14", W14.w14),
            new(XNamespace.Xmlns + "wpg", WPG.wpg),
            new(XNamespace.Xmlns + "wpi", WPI.wpi),
            new(XNamespace.Xmlns + "wne", WNE.wne),
            new(XNamespace.Xmlns + "wps", WPS.wps),
            new(MC.Ignorable, "w14 wp14"),
        };
    }

    public class DocumentBuilderException : Exception
    {
        public DocumentBuilderException(string message) : base(message) { }
    }

    public class DocumentBuilderInternalException : Exception
    {
        public DocumentBuilderInternalException(string message) : base(message) { }
    }
}
