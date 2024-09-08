﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace Clippit
{
    public static class PtUtils
    {
        public static string NormalizeDirName(string dirName)
        {
            var d = dirName.Replace('\\', '/');
            if (d[dirName.Length - 1] != '/' && d[dirName.Length - 1] != '\\')
                return d + "/";

            return d;
        }

        public static string MakeValidXml(string p)
        {
            return p.Any(c => c < 0x20)
                ? p.Select(c => c < 0x20 ? $"_{(int)c:X}_" : c.ToString()).StringConcatenate()
                : p;
        }

        public static void AddElementIfMissing(XDocument partXDoc, XElement existing, string newElement)
        {
            if (existing != null)
                return;

            var newXElement = XElement.Parse(newElement);
            newXElement.Attributes().Where(a => a.IsNamespaceDeclaration).Remove();
            if (partXDoc.Root != null)
                partXDoc.Root.Add(newXElement);
        }
    }

    public class MhtParser
    {
        public string MimeVersion;
        public string ContentType;
        public MhtParserPart[] Parts;

        public class MhtParserPart
        {
            public string ContentLocation;
            public string ContentTransferEncoding;
            public string ContentType;
            public string CharSet;
            public string Text;
            public byte[] Binary;
        }

        public static MhtParser Parse(string src)
        {
            string mimeVersion = null;
            string contentType = null;
            string boundary = null;

            var lines = src.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            var priambleKeyWords = new[] { "MIME-VERSION:", "CONTENT-TYPE:" };

            var priamble = lines
                .TakeWhile(l =>
                {
                    var s = l.ToUpper();
                    return priambleKeyWords.Any(pk => s.StartsWith(pk));
                })
                .ToArray();

            foreach (var item in priamble)
            {
                if (item.ToUpper().StartsWith("MIME-VERSION:"))
                    mimeVersion = item.Substring("MIME-VERSION:".Length).Trim();
                else if (item.ToUpper().StartsWith("CONTENT-TYPE:"))
                {
                    var contentTypeLine = item.Substring("CONTENT-TYPE:".Length).Trim();
                    var spl = contentTypeLine.Split(';').Select(z => z.Trim()).ToArray();
                    foreach (var s in spl)
                    {
                        if (s.StartsWith("boundary"))
                        {
                            var begText = "boundary=\"";
                            var begLen = begText.Length;
                            boundary = s.Substring(begLen, s.Length - begLen - 1).TrimStart('-');
                            continue;
                        }
                        if (contentType == null)
                        {
                            contentType = s;
                            continue;
                        }
                        throw new OpenXmlPowerToolsException("Unexpected content in MHTML");
                    }
                }
            }

            var grouped = lines
                .Skip(priamble.Length)
                .GroupAdjacent(l =>
                {
                    var b = l.TrimStart('-') == boundary;
                    return b;
                })
                .Where(g => g.Key == false)
                .ToArray();

            var parts = grouped
                .Select(rp =>
                {
                    var partPriambleKeyWords = new[]
                    {
                        "CONTENT-LOCATION:",
                        "CONTENT-TRANSFER-ENCODING:",
                        "CONTENT-TYPE:",
                    };

                    var partPriamble = rp.TakeWhile(l =>
                        {
                            var s = l.ToUpper();
                            return partPriambleKeyWords.Any(pk => s.StartsWith(pk));
                        })
                        .ToArray();

                    string contentLocation = null;
                    string contentTransferEncoding = null;
                    string partContentType = null;
                    string partCharSet = null;
                    byte[] partBinary = null;

                    foreach (var item in partPriamble)
                    {
                        if (item.ToUpper().StartsWith("CONTENT-LOCATION:"))
                            contentLocation = item.Substring("CONTENT-LOCATION:".Length).Trim();
                        else if (item.ToUpper().StartsWith("CONTENT-TRANSFER-ENCODING:"))
                            contentTransferEncoding = item.Substring("CONTENT-TRANSFER-ENCODING:".Length).Trim();
                        else if (item.ToUpper().StartsWith("CONTENT-TYPE:"))
                            partContentType = item.Substring("CONTENT-TYPE:".Length).Trim();
                    }

                    var blankLinesAtBeginning = rp.Skip(partPriamble.Length).TakeWhile(l => l == "").Count();

                    var partText = rp.Skip(partPriamble.Length)
                        .Skip(blankLinesAtBeginning)
                        .Select(l => l + Environment.NewLine)
                        .StringConcatenate();

                    if (partContentType != null && partContentType.Contains(";"))
                    {
                        string thisPartContentType = null;
                        var spl = partContentType.Split(';').Select(s => s.Trim()).ToArray();
                        foreach (var s in spl)
                        {
                            if (s.StartsWith("charset"))
                            {
                                var begText = "charset=\"";
                                var begLen = begText.Length;
                                partCharSet = s.Substring(begLen, s.Length - begLen - 1);
                                continue;
                            }
                            if (thisPartContentType == null)
                            {
                                thisPartContentType = s;
                                continue;
                            }
                            throw new OpenXmlPowerToolsException("Unexpected content in MHTML");
                        }
                        partContentType = thisPartContentType;
                    }

                    if (contentTransferEncoding != null && contentTransferEncoding.ToUpper() == "BASE64")
                    {
                        partBinary = Convert.FromBase64String(partText);
                    }

                    return new MhtParserPart()
                    {
                        ContentLocation = contentLocation,
                        ContentTransferEncoding = contentTransferEncoding,
                        ContentType = partContentType,
                        CharSet = partCharSet,
                        Text = partText,
                        Binary = partBinary,
                    };
                })
                .Where(p => p.ContentType != null)
                .ToArray();

            return new MhtParser()
            {
                ContentType = contentType,
                MimeVersion = mimeVersion,
                Parts = parts,
            };
        }
    }

    public class Normalizer
    {
        public static XDocument Normalize(XDocument source, XmlSchemaSet schema)
        {
            var havePSVI = false;
            // validate, throw errors, add PSVI information
            if (schema != null)
            {
                source.Validate(schema, null, true);
                havePSVI = true;
            }
            return new XDocument(
                source.Declaration,
                source
                    .Nodes()
                    .Select(n =>
                    {
                        return n switch
                        {
                            // Remove comments, processing instructions, and text nodes that are
                            // children of XDocument.  Only white space text nodes are allowed as
                            // children of a document, so we can remove all text nodes.
                            XComment or XProcessingInstruction or XText => null,
                            XElement e => NormalizeElement(e, havePSVI),
                            _ => n,
                        };
                    })
            );
        }

        public static bool DeepEqualsWithNormalization(XDocument doc1, XDocument doc2, XmlSchemaSet schemaSet)
        {
            var d1 = Normalize(doc1, schemaSet);
            var d2 = Normalize(doc2, schemaSet);
            return XNode.DeepEquals(d1, d2);
        }

        private static IEnumerable<XAttribute> NormalizeAttributes(XElement element, bool havePSVI)
        {
            return element
                .Attributes()
                .Where(a =>
                    !a.IsNamespaceDeclaration && a.Name != Xsi.schemaLocation && a.Name != Xsi.noNamespaceSchemaLocation
                )
                .OrderBy(a => a.Name.NamespaceName)
                .ThenBy(a => a.Name.LocalName)
                .Select(a =>
                {
                    if (havePSVI)
                    {
                        var dt = a.GetSchemaInfo().SchemaType.TypeCode;
                        switch (dt)
                        {
                            case XmlTypeCode.Boolean:
                                return new XAttribute(a.Name, (bool)a);
                            case XmlTypeCode.DateTime:
                                return new XAttribute(a.Name, (DateTime)a);
                            case XmlTypeCode.Decimal:
                                return new XAttribute(a.Name, (decimal)a);
                            case XmlTypeCode.Double:
                                return new XAttribute(a.Name, (double)a);
                            case XmlTypeCode.Float:
                                return new XAttribute(a.Name, (float)a);
                            case XmlTypeCode.HexBinary:
                            case XmlTypeCode.Language:
                                return new XAttribute(a.Name, ((string)a).ToLower());
                        }
                    }
                    return a;
                });
        }

        private static XNode NormalizeNode(XNode node, bool havePSVI)
        {
            return node switch
            {
                // trim comments and processing instructions from normalized tree
                XComment or XProcessingInstruction => null,
                XElement e => NormalizeElement(e, havePSVI),
                _ => node,
            };
            // Only thing left is XCData and XText, so clone them
        }

        private static XElement NormalizeElement(XElement element, bool havePSVI)
        {
            if (havePSVI)
            {
                var dt = element.GetSchemaInfo();
                switch (dt.SchemaType.TypeCode)
                {
                    case XmlTypeCode.Boolean:
                        return new XElement(element.Name, NormalizeAttributes(element, havePSVI), (bool)element);
                    case XmlTypeCode.DateTime:
                        return new XElement(element.Name, NormalizeAttributes(element, havePSVI), (DateTime)element);
                    case XmlTypeCode.Decimal:
                        return new XElement(element.Name, NormalizeAttributes(element, havePSVI), (decimal)element);
                    case XmlTypeCode.Double:
                        return new XElement(element.Name, NormalizeAttributes(element, havePSVI), (double)element);
                    case XmlTypeCode.Float:
                        return new XElement(element.Name, NormalizeAttributes(element, havePSVI), (float)element);
                    case XmlTypeCode.HexBinary:
                    case XmlTypeCode.Language:
                        return new XElement(
                            element.Name,
                            NormalizeAttributes(element, havePSVI),
                            ((string)element).ToLower()
                        );
                    default:
                        return new XElement(
                            element.Name,
                            NormalizeAttributes(element, havePSVI),
                            element.Nodes().Select(n => NormalizeNode(n, havePSVI))
                        );
                }
            }
            else
            {
                return new XElement(
                    element.Name,
                    NormalizeAttributes(element, havePSVI),
                    element.Nodes().Select(n => NormalizeNode(n, havePSVI))
                );
            }
        }
    }

    public class FileUtils
    {
        public static DirectoryInfo GetDateTimeStampedDirectoryInfo(string prefix)
        {
            var now = DateTime.Now;
            var dirName =
                prefix
                + $"-{now.Year - 2000:00}-{now.Month:00}-{now.Day:00}-{now.Hour:00}{now.Minute:00}{now.Second:00}";
            return new DirectoryInfo(dirName);
        }

        public static FileInfo GetDateTimeStampedFileInfo(string prefix, string suffix)
        {
            var now = DateTime.Now;
            var fileName =
                prefix
                + $"-{now.Year - 2000:00}-{now.Month:00}-{now.Day:00}-{now.Hour:00}{now.Minute:00}{now.Second:00}"
                + suffix;
            return new FileInfo(fileName);
        }

        public static void ThreadSafeCreateDirectory(DirectoryInfo dir)
        {
            while (true)
            {
                if (dir.Exists)
                    break;

                try
                {
                    dir.Create();
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        public static void ThreadSafeCopy(FileInfo sourceFile, FileInfo destFile)
        {
            while (true)
            {
                if (destFile.Exists)
                    break;

                try
                {
                    File.Copy(sourceFile.FullName, destFile.FullName);
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        public static void ThreadSafeCreateEmptyTextFileIfNotExist(FileInfo file)
        {
            while (true)
            {
                if (file.Exists)
                    break;

                try
                {
                    File.WriteAllText(file.FullName, "");
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        internal static void ThreadSafeAppendAllLines(FileInfo file, string[] strings)
        {
            while (true)
            {
                try
                {
                    File.AppendAllLines(file.FullName, strings);
                    break;
                }
                catch (IOException)
                {
                    System.Threading.Thread.Sleep(50);
                }
            }
        }

        public static List<string> GetFilesRecursive(DirectoryInfo dir, string searchPattern)
        {
            var fileList = new List<string>();
            GetFilesRecursiveInternal(dir, searchPattern, fileList);
            return fileList;
        }

        private static void GetFilesRecursiveInternal(DirectoryInfo dir, string searchPattern, List<string> fileList)
        {
            fileList.AddRange(dir.GetFiles(searchPattern).Select(file => file.FullName));
            foreach (var subdir in dir.GetDirectories())
                GetFilesRecursiveInternal(subdir, searchPattern, fileList);
        }

        public static List<string> GetFilesRecursive(DirectoryInfo dir)
        {
            var fileList = new List<string>();
            GetFilesRecursiveInternal(dir, fileList);
            return fileList;
        }

        private static void GetFilesRecursiveInternal(DirectoryInfo dir, List<string> fileList)
        {
            fileList.AddRange(dir.GetFiles().Select(file => file.FullName));
            foreach (var subdir in dir.GetDirectories())
                GetFilesRecursiveInternal(subdir, fileList);
        }
    }

    public static class PtExtensions
    {
        public static XElement GetXElement(this XmlNode node)
        {
            var xDoc = new XDocument();
            using (var xmlWriter = xDoc.CreateWriter())
                node.WriteTo(xmlWriter);
            return xDoc.Root;
        }

        public static XmlNode GetXmlNode(this XElement element)
        {
            var xmlDoc = new XmlDocument();
            using var xmlReader = element.CreateReader();
            xmlDoc.Load(xmlReader);
            return xmlDoc;
        }

        public static XDocument GetXDocument(this XmlDocument document)
        {
            var xDoc = new XDocument();
            using (var xmlWriter = xDoc.CreateWriter())
                document.WriteTo(xmlWriter);

            var decl = document.ChildNodes.OfType<XmlDeclaration>().FirstOrDefault();
            if (decl != null)
                xDoc.Declaration = new XDeclaration(decl.Version, decl.Encoding, decl.Standalone);

            return xDoc;
        }

        public static XmlDocument GetXmlDocument(this XDocument document)
        {
            var xmlDoc = new XmlDocument();
            using var xmlReader = document.CreateReader();
            xmlDoc.Load(xmlReader);
            if (document.Declaration != null)
            {
                var dec = xmlDoc.CreateXmlDeclaration(
                    document.Declaration.Version,
                    document.Declaration.Encoding,
                    document.Declaration.Standalone
                );
                xmlDoc.InsertBefore(dec, xmlDoc.FirstChild);
            }

            return xmlDoc;
        }

        public static string StringConcatenate(this IEnumerable<string> source)
        {
            return source.Aggregate(new StringBuilder(), (sb, s) => sb.Append(s), sb => sb.ToString());
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source, Func<T, string> projectionFunc)
        {
            return source.Aggregate(new StringBuilder(), (sb, i) => sb.Append(projectionFunc(i)), sb => sb.ToString());
        }

        public static IEnumerable<TResult> PtZip<TFirst, TSecond, TResult>(
            this IEnumerable<TFirst> first,
            IEnumerable<TSecond> second,
            Func<TFirst, TSecond, TResult> func
        )
        {
            using var ie1 = first.GetEnumerator();
            using var ie2 = second.GetEnumerator();
            while (ie1.MoveNext() && ie2.MoveNext())
                yield return func(ie1.Current, ie2.Current);
        }

        public static IEnumerable<IGrouping<TKey, TSource>> GroupAdjacent<TSource, TKey>(
            this IEnumerable<TSource> source,
            Func<TSource, TKey> keySelector
        )
        {
            var last = default(TKey);
            var haveLast = false;
            var list = new List<TSource>();

            foreach (var s in source)
            {
                var k = keySelector(s);
                if (haveLast)
                {
                    if (!k.Equals(last))
                    {
                        yield return new GroupOfAdjacent<TSource, TKey>(list, last);

                        list = new List<TSource> { s };
                        last = k;
                    }
                    else
                    {
                        list.Add(s);
                        last = k;
                    }
                }
                else
                {
                    list.Add(s);
                    last = k;
                    haveLast = true;
                }
            }
            if (haveLast)
                yield return new GroupOfAdjacent<TSource, TKey>(list, last);
        }

        private static void InitializeSiblingsReverseDocumentOrder(XElement element)
        {
            XElement prev = null;
            foreach (var e in element.Elements())
            {
                e.AddAnnotation(new SiblingsReverseDocumentOrderInfo { PreviousSibling = prev });
                prev = e;
            }
        }

        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public static IEnumerable<XElement> SiblingsBeforeSelfReverseDocumentOrder(this XElement element)
        {
            if (element.Annotation<SiblingsReverseDocumentOrderInfo>() == null)
                InitializeSiblingsReverseDocumentOrder(element.Parent);
            var current = element;
            while (true)
            {
                var previousElement = current.Annotation<SiblingsReverseDocumentOrderInfo>().PreviousSibling;
                if (previousElement == null)
                    yield break;

                yield return previousElement;

                current = previousElement;
            }
        }

        private static void InitializeDescendantsReverseDocumentOrder(XElement element)
        {
            XElement prev = null;
            foreach (var e in element.Descendants())
            {
                e.AddAnnotation(new DescendantsReverseDocumentOrderInfo { PreviousElement = prev });
                prev = e;
            }
        }

        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public static IEnumerable<XElement> DescendantsBeforeSelfReverseDocumentOrder(this XElement element)
        {
            if (element.Annotation<DescendantsReverseDocumentOrderInfo>() == null)
                InitializeDescendantsReverseDocumentOrder(element.AncestorsAndSelf().Last());
            var current = element;
            while (true)
            {
                var previousElement = current.Annotation<DescendantsReverseDocumentOrderInfo>().PreviousElement;
                if (previousElement == null)
                    yield break;

                yield return previousElement;

                current = previousElement;
            }
        }

        private static void InitializeDescendantsTrimmedReverseDocumentOrder(XElement element, XName trimName)
        {
            XElement prev = null;
            foreach (XElement e in element.DescendantsTrimmed(trimName))
            {
                e.AddAnnotation(new DescendantsTrimmedReverseDocumentOrderInfo { PreviousElement = prev });
                prev = e;
            }
        }

        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public static IEnumerable<XElement> DescendantsTrimmedBeforeSelfReverseDocumentOrder(
            this XElement element,
            XName trimName
        )
        {
            if (element.Annotation<DescendantsTrimmedReverseDocumentOrderInfo>() == null)
            {
                var ances =
                    element.AncestorsAndSelf(W.txbxContent).FirstOrDefault() ?? element.AncestorsAndSelf().Last();
                InitializeDescendantsTrimmedReverseDocumentOrder(ances, trimName);
            }

            var current = element;
            while (true)
            {
                var previousElement = current.Annotation<DescendantsTrimmedReverseDocumentOrderInfo>().PreviousElement;
                if (previousElement == null)
                    yield break;

                yield return previousElement;

                current = previousElement;
            }
        }

        public static string ToStringNewLineOnAttributes(this XElement element)
        {
            var settings = new XmlWriterSettings
            {
                Indent = true,
                OmitXmlDeclaration = true,
                NewLineOnAttributes = true,
            };
            var stringBuilder = new StringBuilder();
            using (var stringWriter = new StringWriter(stringBuilder))
            using (var xmlWriter = XmlWriter.Create(stringWriter, settings))
                element.WriteTo(xmlWriter);
            return stringBuilder.ToString();
        }

        public static IEnumerable<XElement> DescendantsTrimmed(this XElement element, XName trimName)
        {
            return DescendantsTrimmed(element, e => e.Name == trimName);
        }

        public static IEnumerable<XElement> DescendantsTrimmed(this XElement element, Func<XElement, bool> predicate)
        {
            var iteratorStack = new Stack<IEnumerator<XElement>>();
            iteratorStack.Push(element.Elements().GetEnumerator());
            while (iteratorStack.Count > 0)
            {
                while (iteratorStack.Peek().MoveNext())
                {
                    var currentXElement = iteratorStack.Peek().Current;
                    if (predicate(currentXElement))
                    {
                        yield return currentXElement;
                        continue;
                    }
                    yield return currentXElement;
                    iteratorStack.Push(currentXElement.Elements().GetEnumerator());
                }
                iteratorStack.Pop();
            }
        }

        public static IEnumerable<TResult> Rollup<TSource, TResult>(
            this IEnumerable<TSource> source,
            TResult seed,
            Func<TSource, TResult, TResult> projection
        )
        {
            var nextSeed = seed;
            foreach (var src in source)
            {
                var projectedValue = projection(src, nextSeed);
                nextSeed = projectedValue;
                yield return projectedValue;
            }
        }

        public static IEnumerable<TResult> Rollup<TSource, TResult>(
            this IEnumerable<TSource> source,
            TResult seed,
            Func<TSource, TResult, int, TResult> projection
        )
        {
            var nextSeed = seed;
            var index = 0;
            foreach (var src in source)
            {
                var projectedValue = projection(src, nextSeed, index++);
                nextSeed = projectedValue;
                yield return projectedValue;
            }
        }

        public static IEnumerable<TSource> SequenceAt<TSource>(this TSource[] source, int index)
        {
            var i = index;
            while (i < source.Length)
                yield return source[i++];
        }

        public static IEnumerable<T> SkipLast<T>(this IEnumerable<T> source, int count)
        {
            var saveList = new Queue<T>();
            var saved = 0;
            foreach (var item in source)
            {
                if (saved < count)
                {
                    saveList.Enqueue(item);
                    ++saved;
                    continue;
                }

                saveList.Enqueue(item);
                yield return saveList.Dequeue();
            }
        }

        public static bool? ToBoolean(this XAttribute a)
        {
            if (a == null)
                return null;

            var s = ((string)a).ToLower();
            return s switch
            {
                "1" => true,
                "0" => false,
                "true" => true,
                "false" => false,
                "on" => true,
                "off" => false,
                _ => (bool)a,
            };
        }

        private static string GetQName(XElement xe)
        {
            var prefix = xe.GetPrefixOfNamespace(xe.Name.Namespace);
            if (xe.Name.Namespace == XNamespace.None || prefix == null)
                return xe.Name.LocalName;

            return prefix + ":" + xe.Name.LocalName;
        }

        private static string GetQName(XAttribute xa)
        {
            var prefix = xa.Parent != null ? xa.Parent.GetPrefixOfNamespace(xa.Name.Namespace) : null;
            if (xa.Name.Namespace == XNamespace.None || prefix == null)
                return xa.Name.ToString();

            return prefix + ":" + xa.Name.LocalName;
        }

        private static string NameWithPredicate(XElement el)
        {
            if (el.Parent != null && el.Parent.Elements(el.Name).Count() != 1)
                return GetQName(el) + "[" + (el.ElementsBeforeSelf(el.Name).Count() + 1) + "]";
            else
                return GetQName(el);
        }

        public static string StrCat<T>(this IEnumerable<T> source, string separator)
        {
            return source.Aggregate(new StringBuilder(), (sb, i) => sb.Append(i).Append(separator), s => s.ToString());
        }

        public static string GetXPath(this XObject xobj)
        {
            if (xobj.Parent == null)
            {
                return xobj switch
                {
                    XDocument doc => ".",
                    XElement el => "/" + NameWithPredicate(el),
                    XText xt => null,
                    //
                    //the following doesn't work because the XPath data
                    //model doesn't include white space text nodes that
                    //are children of the document.
                    //
                    //return
                    //    "/" +
                    //    (
                    //        xt
                    //        .Document
                    //        .Nodes()
                    //        .OfType<XText>()
                    //        .Count() != 1 ?
                    //        "text()[" +
                    //        (xt
                    //        .NodesBeforeSelf()
                    //        .OfType<XText>()
                    //        .Count() + 1) + "]" :
                    //        "text()"
                    //    );
                    //
                    XComment com when com.Document != null => "/"
                        + (
                            com.Document.Nodes().OfType<XComment>().Count() != 1
                                ? "comment()[" + (com.NodesBeforeSelf().OfType<XComment>().Count() + 1) + "]"
                                : "comment()"
                        ),
                    XProcessingInstruction pi => "/"
                        + (
                            pi.Document != null && pi.Document.Nodes().OfType<XProcessingInstruction>().Count() != 1
                                ? "processing-instruction()["
                                    + (pi.NodesBeforeSelf().OfType<XProcessingInstruction>().Count() + 1)
                                    + "]"
                                : "processing-instruction()"
                        ),
                    _ => null,
                };
            }
            else
            {
                return xobj switch
                {
                    XElement el => "/"
                        + el.Ancestors().InDocumentOrder().Select(e => NameWithPredicate(e)).StrCat("/")
                        + NameWithPredicate(el),
                    XAttribute at when at.Parent != null => "/"
                        + at.Parent.AncestorsAndSelf().InDocumentOrder().Select(e => NameWithPredicate(e)).StrCat("/")
                        + "@"
                        + GetQName(at),
                    XComment com when com.Parent != null => "/"
                        + com.Parent.AncestorsAndSelf().InDocumentOrder().Select(e => NameWithPredicate(e)).StrCat("/")
                        + (
                            com.Parent.Nodes().OfType<XComment>().Count() != 1
                                ? "comment()[" + (com.NodesBeforeSelf().OfType<XComment>().Count() + 1) + "]"
                                : "comment()"
                        ),
                    XCData cd when cd.Parent != null => "/"
                        + cd.Parent.AncestorsAndSelf().InDocumentOrder().Select(e => NameWithPredicate(e)).StrCat("/")
                        + (
                            cd.Parent.Nodes().OfType<XText>().Count() != 1
                                ? "text()[" + (cd.NodesBeforeSelf().OfType<XText>().Count() + 1) + "]"
                                : "text()"
                        ),
                    XText tx when tx.Parent != null => "/"
                        + tx.Parent.AncestorsAndSelf().InDocumentOrder().Select(e => NameWithPredicate(e)).StrCat("/")
                        + (
                            tx.Parent.Nodes().OfType<XText>().Count() != 1
                                ? "text()[" + (tx.NodesBeforeSelf().OfType<XText>().Count() + 1) + "]"
                                : "text()"
                        ),
                    XProcessingInstruction pi when pi.Parent != null => "/"
                        + pi.Parent.AncestorsAndSelf().InDocumentOrder().Select(e => NameWithPredicate(e)).StrCat("/")
                        + (
                            pi.Parent.Nodes().OfType<XProcessingInstruction>().Count() != 1
                                ? "processing-instruction()["
                                    + (pi.NodesBeforeSelf().OfType<XProcessingInstruction>().Count() + 1)
                                    + "]"
                                : "processing-instruction()"
                        ),
                    _ => null,
                };
            }
        }
    }

    public class ExecutableRunner
    {
        public class RunResults
        {
            public int ExitCode;
            public Exception RunException;
            public StringBuilder Output;
            public StringBuilder Error;
        }

        public static RunResults RunExecutable(string executablePath, string arguments, string workingDirectory)
        {
            var runResults = new RunResults
            {
                Output = new StringBuilder(),
                Error = new StringBuilder(),
                RunException = null,
            };
            try
            {
                if (File.Exists(executablePath))
                {
                    using var proc = new Process();
                    proc.StartInfo.FileName = executablePath;
                    proc.StartInfo.Arguments = arguments;
                    proc.StartInfo.WorkingDirectory = workingDirectory;
                    proc.StartInfo.UseShellExecute = false;
                    proc.StartInfo.RedirectStandardOutput = true;
                    proc.StartInfo.RedirectStandardError = true;
                    proc.OutputDataReceived += (o, e) => runResults.Output.Append(e.Data).Append(Environment.NewLine);
                    proc.ErrorDataReceived += (o, e) => runResults.Error.Append(e.Data).Append(Environment.NewLine);
                    proc.Start();
                    proc.BeginOutputReadLine();
                    proc.BeginErrorReadLine();
                    proc.WaitForExit();
                    runResults.ExitCode = proc.ExitCode;
                }
                else
                {
                    throw new ArgumentException("Invalid executable path.", nameof(executablePath));
                }
            }
            catch (Exception e)
            {
                runResults.RunException = e;
            }
            return runResults;
        }
    }

    public class SiblingsReverseDocumentOrderInfo
    {
        public XElement PreviousSibling;
    }

    public class DescendantsReverseDocumentOrderInfo
    {
        public XElement PreviousElement;
    }

    public class DescendantsTrimmedReverseDocumentOrderInfo
    {
        public XElement PreviousElement;
    }

    public class GroupOfAdjacent<TSource, TKey> : IGrouping<TKey, TSource>
    {
        public GroupOfAdjacent(List<TSource> source, TKey key)
        {
            GroupList = source;
            Key = key;
        }

        public TKey Key { get; set; }
        private List<TSource> GroupList { get; set; }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<TSource>)this).GetEnumerator();
        }

        IEnumerator<TSource> IEnumerable<TSource>.GetEnumerator()
        {
            return ((IEnumerable<TSource>)GroupList).GetEnumerator();
        }
    }

    public static class PtBucketTimer
    {
        private class BucketInfo
        {
            public int Count;
            public TimeSpan Time;
        }

        private static string LastBucket;
        private static DateTime LastTime;
        private static Dictionary<string, BucketInfo> Buckets;

        public static void Bucket(string bucket)
        {
            var now = DateTime.Now;
            if (LastBucket != null)
            {
                var d = now - LastTime;
                if (Buckets.ContainsKey(LastBucket))
                {
                    Buckets[LastBucket].Count = Buckets[LastBucket].Count + 1;
                    Buckets[LastBucket].Time += d;
                }
                else
                {
                    Buckets.Add(LastBucket, new BucketInfo() { Count = 1, Time = d });
                }
            }
            LastBucket = bucket;
            LastTime = now;
        }

        public static string DumpBucketsByKey()
        {
            var sb = new StringBuilder();
            foreach (var bucket in Buckets.OrderBy(b => b.Key))
            {
                var ts = bucket.Value.Time.ToString();
                if (ts.Contains('.'))
                    ts = ts.Substring(0, ts.Length - 5);
                var s = bucket.Key.PadRight(60, '-') + "  " + $"{bucket.Value.Count:00000000}" + "  " + ts;
                sb.Append(s + Environment.NewLine);
            }
            var total = Buckets.Aggregate(TimeSpan.Zero, (t, b) => t + b.Value.Time);
            var tz = total.ToString();
            sb.Append($"Total: {tz.Substring(0, tz.Length - 5)}");
            return sb.ToString();
        }

        public static string DumpBucketsByTime()
        {
            var sb = new StringBuilder();
            foreach (var bucket in Buckets.OrderBy(b => b.Value.Time))
            {
                var ts = bucket.Value.Time.ToString();
                if (ts.Contains('.'))
                    ts = ts.Substring(0, ts.Length - 5);
                var s = bucket.Key.PadRight(60, '-') + "  " + $"{bucket.Value.Count:00000000}" + "  " + ts;
                sb.Append(s + Environment.NewLine);
            }
            var total = Buckets.Aggregate(TimeSpan.Zero, (t, b) => t + b.Value.Time);
            var tz = total.ToString();
            sb.Append($"Total: {tz.Substring(0, tz.Length - 5)}");
            return sb.ToString();
        }

        public static void Init()
        {
            Buckets = new Dictionary<string, BucketInfo>();
        }
    }

    public class XEntity : XText
    {
        public override void WriteTo(XmlWriter writer)
        {
            if (Value.Substring(0, 1) == "#")
            {
                writer.WriteRaw($"&{Value};");
            }
            else
                writer.WriteEntityRef(Value);
        }

        public XEntity(string value)
            : base(value) { }
    }

    public static class Xsi
    {
        public static XNamespace xsi = "http://www.w3.org/2001/XMLSchema-instance";

        public static XName schemaLocation = xsi + "schemaLocation";
        public static XName noNamespaceSchemaLocation = xsi + "noNamespaceSchemaLocation";
    }
}
