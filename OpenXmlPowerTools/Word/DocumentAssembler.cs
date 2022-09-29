// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;
using System.Xml.XPath;
using DocumentFormat.OpenXml.Packaging;
using Path = System.IO.Path;

namespace Clippit.Word
{
    public static class DocumentAssembler
    {
        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XmlDocument data, out bool templateError)
        {
            var xDoc = data.GetXDocument();
            return AssembleDocument(templateDoc, xDoc.Root, out templateError);
        }

        public static WmlDocument AssembleDocument(WmlDocument templateDoc, XElement data, out bool templateError)
        {
            var byteArray = templateDoc.DocumentByteArray;
            using var mem = new MemoryStream();
            mem.Write(byteArray, 0, byteArray.Length);

            using (var wordDoc = WordprocessingDocument.Open(mem, true))
            {
                if (RevisionAccepter.HasTrackedRevisions(wordDoc))
                    throw new OpenXmlPowerToolsException("Invalid DocumentAssembler template - contains tracked revisions");

                // calculate and store the max docPr id for later use when adding image objects
                var macDocPrId = GetMaxDocPrId(wordDoc);

                var te = new TemplateError();
                foreach (var part in wordDoc.ContentParts())
                {
                    ProcessTemplatePart(data, te, part);
                }
                templateError = te.HasError;

                // update image docPr ids for the whole document
                FixUpDocPrIds(wordDoc, macDocPrId);
            }

            var assembledDocument = new WmlDocument("TempFileName.docx", mem.ToArray());
            return assembledDocument;
        }

        private static void ProcessTemplatePart(XElement data, TemplateError te, OpenXmlPart part)
        {
            var xDoc = part.GetXDocument();

            var xDocRoot = RemoveGoBackBookmarks(xDoc.Root);

            // process diagrams part
            // TODO: consider splitting this method into two for clarity, so, for example, there will be
            // TODO: pipeline-like processing: first the diagram, then the document part, or vice versa
            var diagramPart = part.GetPartsOfType<DiagramDataPart>().FirstOrDefault();
            if (diagramPart != null)
            {
                var diagramDoc = diagramPart.GetXDocument();
                if (diagramDoc != null)
                {
                    var dataPartRoot = diagramDoc.Root;
                    if (dataPartRoot != null)
                    {
                        dataPartRoot = (XElement)TransformToMetadata(dataPartRoot, te);
                        // do the actual content replacement
                        dataPartRoot = (XElement)ContentReplacementTransform(dataPartRoot, data, te, part);
                        diagramDoc.Elements().First().ReplaceWith(dataPartRoot);
                        diagramPart.PutXDocument();
                    }
                }
            }

            // content controls in cells can surround the W.tc element, so transform so that such content controls are within the cell content
            xDocRoot = (XElement)NormalizeContentControlsInCells(xDocRoot);

            xDocRoot = (XElement)TransformToMetadata(xDocRoot, te);

            // Table might have been placed at run-level, when it should be at block-level, so fix this.
            // Repeat, EndRepeat, Conditional, EndConditional are allowed at run level, but only if there is a matching pair
            // if there is only one Repeat, EndRepeat, Conditional, EndConditional, then move to block level.
            // if there is a matching pair, then is OK.
            xDocRoot = (XElement)ForceBlockLevelAsAppropriate(xDocRoot, te);

            NormalizeTablesRepeatAndConditional(xDocRoot, te);

            // any EndRepeat, EndConditional that remain are orphans, so replace with an error
            ProcessOrphanEndRepeatEndConditional(xDocRoot, te);

            // do the actual content replacement
            xDocRoot = (XElement)ContentReplacementTransform(xDocRoot, data, te, part);

            xDoc.Elements().First().ReplaceWith(xDocRoot);
            part.PutXDocument();
        }

        private static readonly XName[] s_metaToForceToBlock = {
            PA.Conditional,
            PA.EndConditional,
            PA.Repeat,
            PA.EndRepeat,
            PA.Table,
            PA.Image
        };

        private static object ForceBlockLevelAsAppropriate(XNode node, TemplateError te)
        {
            if (node is not XElement element)
                return node;

            if (element.Name == W.p)
            {
                var childMeta = element.Elements().Where(n => s_metaToForceToBlock.Contains(n.Name)).ToList();
                if (childMeta.Count() == 1)
                {
                    var child = childMeta.First();
                    var otherTextInParagraph = element.Elements(W.r).Elements(W.t).Select(t => (string)t).StringConcatenate().Trim();
                    if (otherTextInParagraph != "")
                    {
                        var newPara = new XElement(element);
                        var newMeta = newPara.Elements().First(n => s_metaToForceToBlock.Contains(n.Name));
                        newMeta.ReplaceWith(CreateRunErrorMessage("Error: Unmatched metadata can't be in paragraph with other text", te));
                        return newPara;
                    }
                    var meta = new XElement(child.Name,
                        child.Attributes(),
                        new XElement(W.p,
                            element.Attributes(),
                            element.Elements(W.pPr),
                            child.Elements()));
                    return meta;
                }
                var count = childMeta.Count();
                if (count % 2 == 0)
                {
                    if (childMeta.Count(c => c.Name == PA.Repeat) != childMeta.Count(c => c.Name == PA.EndRepeat))
                        return CreateContextErrorMessage(element, "Error: Mismatch Repeat / EndRepeat at run level", te);
                    if (childMeta.Count(c => c.Name == PA.Conditional) != childMeta.Count(c => c.Name == PA.EndConditional))
                        return CreateContextErrorMessage(element, "Error: Mismatch Conditional / EndConditional at run level", te);
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => ForceBlockLevelAsAppropriate(n, te)));
                }
                else
                {
                    return CreateContextErrorMessage(element, "Error: Invalid metadata at run level", te);
                }
            }
            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => ForceBlockLevelAsAppropriate(n, te)));
        }

        private static void ProcessOrphanEndRepeatEndConditional(XElement xDocRoot, TemplateError te)
        {
            foreach (var element in xDocRoot.Descendants(PA.EndRepeat).ToList())
            {
                var error = CreateContextErrorMessage(element, "Error: EndRepeat without matching Repeat", te);
                element.ReplaceWith(error);
            }
            foreach (var element in xDocRoot.Descendants(PA.EndConditional).ToList())
            {
                var error = CreateContextErrorMessage(element, "Error: EndConditional without matching Conditional", te);
                element.ReplaceWith(error);
            }
        }

        private static XElement RemoveGoBackBookmarks(XElement xElement)
        {
            var cloneXDoc = new XElement(xElement);
            while (true)
            {
                var bm = cloneXDoc.DescendantsAndSelf(W.bookmarkStart).FirstOrDefault(b => (string)b.Attribute(W.name) == "_GoBack");
                if (bm is null)
                    break;
                var id = (string)bm.Attribute(W.id);
                var endBm = cloneXDoc.DescendantsAndSelf(W.bookmarkEnd).FirstOrDefault(b => (string)b.Attribute(W.id) == id);
                bm.Remove();
                endBm?.Remove();
            }
            return cloneXDoc;
        }

        // this transform inverts content controls that surround W.tc elements.  After transforming, the W.tc will contain
        // the content control, which contains the paragraph content of the cell.
        private static object NormalizeContentControlsInCells(XNode node)
        {
            if (node is not XElement element)
                return node;

            if (element.Name == W.sdt && element.Parent.Name == W.tr)
            {
                var newCell = new XElement(W.tc,
                    element.Elements(W.tc).Elements(W.tcPr),
                    new XElement(W.sdt,
                        element.Elements(W.sdtPr),
                        element.Elements(W.sdtEndPr),
                        new XElement(W.sdtContent,
                            element.Elements(W.sdtContent).Elements(W.tc).Elements().Where(e => e.Name != W.tcPr))));
                return newCell;
            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(NormalizeContentControlsInCells));
        }

        // The following method is written using tree modification, not RPFT, because it is easier to write in this fashion.
        // These types of operations are not as easy to write using RPFT.
        // Unless you are completely clear on the semantics of LINQ to XML DML, do not make modifications to this method.
        private static void NormalizeTablesRepeatAndConditional(XElement xDoc, TemplateError te)
        {
            var tables = xDoc.Descendants(PA.Table).ToList();
            foreach (var table in tables)
            {
                var followingElement = table.ElementsAfterSelf().FirstOrDefault(e => e.Name == W.tbl || e.Name == W.p);
                if (followingElement == null || followingElement.Name != W.tbl)
                {
                    table.ReplaceWith(CreateParaErrorMessage("Table metadata is not immediately followed by a table", te));
                    continue;
                }
                // remove superfluous paragraph from Table metadata
                table.RemoveNodes();
                // detach w:tbl from parent, and add to Table metadata
                followingElement.Remove();
                table.Add(followingElement);
            }

            var images = xDoc.Descendants(PA.Image).ToList();
            foreach (var image in images)
            {
                var followingElement = image.ElementsAfterSelf().FirstOrDefault(e => e.Name == W.sdt || e.Name == W.p);

                if (followingElement == null)
                {
                    image.ReplaceWith(CreateParaErrorMessage("Image metadata is not immediately followed by an image", te));
                    continue;
                }

                // get sdt element (can also be within a paragraph) and check it's contents
                var sdt = followingElement.Name == W.p ? followingElement.Elements().FirstOrDefault(e => e.Name == W.sdt) : followingElement;

                if (sdt != null && sdt.Name == W.sdt)
                {
                    // get sdt properties
                    var sdtPr = sdt.Elements().FirstOrDefault(e => e.Name == W.sdtPr);
                    if (sdtPr != null)
                    {
                        // check for properties if contain picture
                        var picture = sdtPr.Elements().FirstOrDefault(e => e.Name == W.picture);
                        if (picture == null)
                        {
                            image.ReplaceWith(
                                CreateParaErrorMessage("Image metadata does not contain picture element", te));
                            continue;
                        }
                    }
                }
                else
                {
                    // there might be the image without surrounding content control
                    image.RemoveNodes();
                    followingElement.Remove();
                    image.Add(followingElement);
                    continue;
                }

                // remove superfluous paragraph from Image metadata
                image.RemoveNodes();
                // detach w:sdt from parent, and add to Image metadata
                followingElement.Remove();
                image.Add(followingElement);
            }

            var repeatDepth = 0;
            var conditionalDepth = 0;
            foreach (var metadata in xDoc.Descendants().Where(d =>
                    d.Name == PA.Repeat ||
                    d.Name == PA.Conditional ||
                    d.Name == PA.EndRepeat ||
                    d.Name == PA.EndConditional))
            {
                if (metadata.Name == PA.Repeat)
                {
                    ++repeatDepth;
                    metadata.Add(new XAttribute(PA.Depth, repeatDepth));
                }
                else if (metadata.Name == PA.EndRepeat)
                {
                    metadata.Add(new XAttribute(PA.Depth, repeatDepth));
                    --repeatDepth;
                }
                else if (metadata.Name == PA.Conditional)
                {
                    ++conditionalDepth;
                    metadata.Add(new XAttribute(PA.Depth, conditionalDepth));
                }
                else if (metadata.Name == PA.EndConditional)
                {
                    metadata.Add(new XAttribute(PA.Depth, conditionalDepth));
                    --conditionalDepth;
                }
            }

            while (true)
            {
                var didReplace = false;
                foreach (var metadata in xDoc.Descendants().Where(d => (d.Name == PA.Repeat || d.Name == PA.Conditional || d.Name == PA.Image) && d.Attribute(PA.Depth) != null).ToList())
                {
                    var depth = (int)metadata.Attribute(PA.Depth);
                    XName matchingEndName = null;
                    if (metadata.Name == PA.Repeat)
                        matchingEndName = PA.EndRepeat;
                    else if (metadata.Name == PA.Conditional)
                        matchingEndName = PA.EndConditional;
                    if (matchingEndName == null)
                        throw new OpenXmlPowerToolsException("Internal error");
                    var matchingEnd = metadata.ElementsAfterSelf(matchingEndName).FirstOrDefault(end => (int)end.Attribute(PA.Depth) == depth);
                    if (matchingEnd == null)
                    {
                        metadata.ReplaceWith(CreateParaErrorMessage($"{metadata.Name.LocalName} does not have matching {matchingEndName.LocalName}", te));
                        continue;
                    }
                    metadata.RemoveNodes();
                    var contentBetween = metadata.ElementsAfterSelf().TakeWhile(after => after != matchingEnd).ToList();
                    foreach (var item in contentBetween)
                        item.Remove();
                    contentBetween = contentBetween.Where(n => n.Name != W.bookmarkStart && n.Name != W.bookmarkEnd).ToList();
                    metadata.Add(contentBetween);
                    metadata.Attributes(PA.Depth).Remove();
                    matchingEnd.Remove();
                    didReplace = true;
                    break;
                }
                if (!didReplace)
                    break;
            }
        }

        private static readonly List<string> s_aliasList = new()
        {
            "Image",
            "Content",
            "Table",
            "Repeat",
            "EndRepeat",
            "Conditional",
            "EndConditional",
        };

        private static object TransformToMetadata(XNode node, TemplateError te)
        {
            if (node is not XElement element)
                return node;

            if (element.Name == W.sdt)
            {
                var alias = (string)element.Elements(W.sdtPr).Elements(W.alias).Attributes(W.val).FirstOrDefault();
                if (string.IsNullOrEmpty(alias) || s_aliasList.Contains(alias))
                {
                    var ccContents = element
                        .DescendantsTrimmed(W.txbxContent)
                        .Where(e => e.Name == W.t)
                        .Select(t => (string)t)
                        .StringConcatenate()
                        .Trim()
                        .Replace('“', '"')
                        .Replace('”', '"');
                    if (ccContents.StartsWith("<"))
                    {
                        var xml = TransformXmlTextToMetadata(te, ccContents);
                        if (xml.Name == W.p || xml.Name == W.r)  // this means there was an error processing the XML.
                        {
                            if (element.Parent.Name == W.p)
                                return xml.Elements(W.r);
                            return xml;
                        }
                        if (alias != null && xml.Name.LocalName != alias)
                        {
                            return element.Parent.Name == W.p
                                ? CreateRunErrorMessage("Error: Content control alias does not match metadata element name", te)
                                : CreateParaErrorMessage("Error: Content control alias does not match metadata element name", te);
                        }
                        xml.Add(element.Elements(W.sdtContent).Elements());
                        return xml;
                    }
                    return new XElement(element.Name,
                        element.Attributes(),
                        element.Nodes().Select(n => TransformToMetadata(n, te)));
                }
                return new XElement(element.Name,
                    element.Attributes(),
                    element.Nodes().Select(n => TransformToMetadata(n, te)));
            }
            if (element.Name == A.r)
            {
                var paraContents = element
                    .DescendantsTrimmed(W.txbxContent)
                    .Where(e => e.Name == A.t)
                    .Select(t => (string)t)
                    .StringConcatenate()
                    .Trim();
                var occurrences = paraContents.Select((_, i) => paraContents.Substring(i)).Count(sub => sub.StartsWith("<#"));
                if (paraContents.StartsWith("<#") && paraContents.EndsWith("#>") && occurrences == 1)
                {
                    var xmlText = paraContents.Substring(2, paraContents.Length - 4).Trim();
                    var xml = TransformXmlTextToMetadata(te, xmlText);
                    if (xml.Name == W.p || xml.Name == W.r)
                        return xml;
                    xml.Add(element);
                    return xml;
                }
                if (paraContents.Contains("<#"))
                {
                    var runReplacementInfo = new List<RunReplacementInfo>();
                    var thisGuid = Guid.NewGuid().ToString();
                    var r = new Regex("<#.*?#>");
                    XElement xml;
                    OpenXmlRegex.Replace(new[] { element }, r, thisGuid, (_, match) =>
                    {
                        var matchString = match.Value.Trim();
                        var xmlText = matchString.Substring(2, matchString.Length - 4).Trim().Replace('“', '"').Replace('”', '"');
                        try
                        {
                            xml = XElement.Parse(xmlText);
                        }
                        catch (XmlException e)
                        {
                            var rri = new RunReplacementInfo
                            {
                                Xml = null,
                                XmlExceptionMessage = "XmlException: " + e.Message,
                                SchemaValidationMessage = null,
                            };
                            runReplacementInfo.Add(rri);
                            return true;
                        }
                        var schemaError = ValidatePerSchema(xml);
                        if (schemaError != null)
                        {
                            var rri = new RunReplacementInfo
                            {
                                Xml = null,
                                XmlExceptionMessage = null,
                                SchemaValidationMessage = "Schema Validation Error: " + schemaError,
                            };
                            runReplacementInfo.Add(rri);
                            return true;
                        }
                        var rri2 = new RunReplacementInfo
                        {
                            Xml = xml,
                            XmlExceptionMessage = null,
                            SchemaValidationMessage = null,
                        };
                        runReplacementInfo.Add(rri2);
                        return true;
                    }, false);

                    var newPara = new XElement(element);
                    foreach (var rri in runReplacementInfo)
                    {
                        var runToReplace = newPara.Descendants(W.r).FirstOrDefault(rn => rn.Value == thisGuid && rn.Parent.Name != PA.Content);
                        if (runToReplace == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        if (rri.XmlExceptionMessage != null)
                            runToReplace.ReplaceWith(CreateRunErrorMessage(rri.XmlExceptionMessage, te));
                        else if (rri.SchemaValidationMessage != null)
                            runToReplace.ReplaceWith(CreateRunErrorMessage(rri.SchemaValidationMessage, te));
                        else
                        {
                            var newXml = new XElement(rri.Xml);
                            newXml.Add(runToReplace);
                            runToReplace.ReplaceWith(newXml);
                        }
                    }
                    var coalescedParagraph = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newPara);
                    return coalescedParagraph;
                }
            }
            if (element.Name == W.p)
            {
                var paraContents = element
                    .DescendantsTrimmed(W.txbxContent)
                    .Where(e => e.Name == W.t)
                    .Select(t => (string)t)
                    .StringConcatenate()
                    .Trim();
                var occurrences = paraContents.Select((_, i) => paraContents.Substring(i)).Count(sub => sub.StartsWith("<#"));
                if (paraContents.StartsWith("<#") && paraContents.EndsWith("#>") && occurrences == 1)
                {
                    var xmlText = paraContents.Substring(2, paraContents.Length - 4).Trim();
                    var xml = TransformXmlTextToMetadata(te, xmlText);
                    if (xml.Name == W.p || xml.Name == W.r)
                        return xml;
                    xml.Add(element);
                    return xml;
                }
                if (paraContents.Contains("<#"))
                {
                    var runReplacementInfo = new List<RunReplacementInfo>();
                    var thisGuid = Guid.NewGuid().ToString();
                    var r = new Regex("<#.*?#>");
                    XElement xml;
                    OpenXmlRegex.Replace(new[] { element }, r, thisGuid, (_, match) =>
                    {
                        var matchString = match.Value.Trim();
                        var xmlText = matchString.Substring(2, matchString.Length - 4).Trim()
                            .Replace('“', '"').Replace('”', '"');

                        try
                        {
                            xml = XElement.Parse(xmlText);
                        }
                        catch (XmlException e)
                        {
                            var rri = new RunReplacementInfo
                            {
                                Xml = null,
                                XmlExceptionMessage = "XmlException: " + e.Message,
                                SchemaValidationMessage = null,
                            };
                            runReplacementInfo.Add(rri);
                            return true;
                        }
                        var schemaError = ValidatePerSchema(xml);
                        if (schemaError != null)
                        {
                            var rri = new RunReplacementInfo
                            {
                                Xml = null,
                                XmlExceptionMessage = null,
                                SchemaValidationMessage = "Schema Validation Error: " + schemaError,
                            };
                            runReplacementInfo.Add(rri);
                            return true;
                        }
                        var rri2 = new RunReplacementInfo
                        {
                            Xml = xml,
                            XmlExceptionMessage = null,
                            SchemaValidationMessage = null,
                        };
                        runReplacementInfo.Add(rri2);
                        return true;
                    }, false);

                    var newPara = new XElement(element);
                    foreach (var rri in runReplacementInfo)
                    {
                        var runToReplace = newPara.Descendants(W.r).FirstOrDefault(rn => rn.Value == thisGuid && rn.Parent.Name != PA.Content);
                        if (runToReplace == null)
                            throw new OpenXmlPowerToolsException("Internal error");
                        if (rri.XmlExceptionMessage != null)
                            runToReplace.ReplaceWith(CreateRunErrorMessage(rri.XmlExceptionMessage, te));
                        else if (rri.SchemaValidationMessage != null)
                            runToReplace.ReplaceWith(CreateRunErrorMessage(rri.SchemaValidationMessage, te));
                        else
                        {
                            var newXml = new XElement(rri.Xml);
                            newXml.Add(runToReplace);
                            runToReplace.ReplaceWith(newXml);
                        }
                    }
                    var coalescedParagraph = WordprocessingMLUtil.CoalesceAdjacentRunsWithIdenticalFormatting(newPara);
                    return coalescedParagraph;
                }
            }

            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => TransformToMetadata(n, te)));
        }

        private static XElement TransformXmlTextToMetadata(TemplateError te, string xmlText)
        {
            XElement xml;
            try
            {
                xml = XElement.Parse(xmlText);
            }
            catch (XmlException e)
            {
                return CreateParaErrorMessage("XmlException: " + e.Message, te);
            }
            var schemaError = ValidatePerSchema(xml);
            if (schemaError is not null)
                return CreateParaErrorMessage("Schema Validation Error: " + schemaError, te);
            return xml;
        }

        private class RunReplacementInfo
        {
            public XElement Xml { get; set; }
            public string XmlExceptionMessage { get; set; }
            public string SchemaValidationMessage { get; set; }
        }

        private static string ValidatePerSchema(XElement element)
        {
            if (s_paSchemaSets == null)
            {
                s_paSchemaSets = new Dictionary<XName, PASchemaSet>
                {
                    {
                        PA.Content,
                        new PASchemaSet
                        {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Content'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Table,
                        new PASchemaSet
                        {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Table'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Repeat,
                        new PASchemaSet
                        {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Repeat'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Optional' type='xs:boolean' use='optional' />
                                      <xs:attribute name='Align' type='xs:string' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.EndRepeat,
                        new PASchemaSet
                        {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='EndRepeat' />
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Conditional,
                        new PASchemaSet
                        {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Conditional'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                      <xs:attribute name='Match' type='xs:string' use='optional' />
                                      <xs:attribute name='NotMatch' type='xs:string' use='optional' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    },
                    {
                        PA.EndConditional,
                        new PASchemaSet
                        {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='EndConditional' />
                                </xs:schema>",
                        }
                    },
                    {
                        PA.Image,
                        new PASchemaSet
                        {
                            XsdMarkup =
                              @"<xs:schema attributeFormDefault='unqualified' elementFormDefault='qualified' xmlns:xs='http://www.w3.org/2001/XMLSchema'>
                                  <xs:element name='Image'>
                                    <xs:complexType>
                                      <xs:attribute name='Select' type='xs:string' use='required' />
                                    </xs:complexType>
                                  </xs:element>
                                </xs:schema>",
                        }
                    }
                };
                foreach (var item in s_paSchemaSets)
                {
                    var itemPAss = item.Value;
                    var schemas = new XmlSchemaSet();
                    schemas.Add("", XmlReader.Create(new StringReader(itemPAss.XsdMarkup)));
                    itemPAss.SchemaSet = schemas;
                }
            }
            if (!s_paSchemaSets.ContainsKey(element.Name))
            {
                return $"Invalid XML: {element.Name.LocalName} is not a valid element";
            }
            var paSchemaSet = s_paSchemaSets[element.Name];
            var d = new XDocument(element);
            string message = null;
            d.Validate(paSchemaSet.SchemaSet, (_, e) =>
            {
                message ??= e.Message;
            }, true);
            return message;
        }

        private static class PA
        {
            public static readonly XName Image = "Image";
            public static readonly XName Content = "Content";
            public static readonly XName Table = "Table";
            public static readonly XName Repeat = "Repeat";
            public static readonly XName EndRepeat = "EndRepeat";
            public static readonly XName Conditional = "Conditional";
            public static readonly XName EndConditional = "EndConditional";

            public static readonly XName Select = "Select";
            public static readonly XName Optional = "Optional";
            public static readonly XName Match = "Match";
            public static readonly XName NotMatch = "NotMatch";
            public static readonly XName Depth = "Depth";
            public static readonly XName Align = "Align";
        }

        private class PASchemaSet
        {
            public string XsdMarkup { get; set; }
            public XmlSchemaSet SchemaSet { get; set; }
        }

        private static Dictionary<XName, PASchemaSet> s_paSchemaSets;

        private class TemplateError
        {
            public bool HasError { get; set; }
        }

        /// <summary>
        /// Gets the next image relationship identifier of given part. The
        /// parts can be either header, footer or main document part. The method
        /// scans for already present relationship identifiers, then increments and
        /// returns the next available value.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <returns>System.String.</returns>
        private static string GetNextImageRelationshipId(OpenXmlPart part)
        {
            switch (part)
            {
                case MainDocumentPart mainDocumentPart:
                    {
                        var imageId = mainDocumentPart.Parts
                            .Select(p => Regex.Match(p.RelationshipId, @"rId(?<rId>\d+)").Groups["rId"].Value)
                            .Max(Convert.ToDecimal);

                        return $"rId{++imageId}";
                    }
                case HeaderPart headerPart:
                    {
                        var imageId = headerPart.Parts
                            .Select(p => Regex.Match(p.RelationshipId, @"rId(?<rId>\d+)").Groups["rId"].Value)
                            .Max(Convert.ToDecimal);

                        return $"rId{++imageId}";
                    }
                case FooterPart footerPart:
                    {
                        var imageId = footerPart.Parts
                            .Select(p => Regex.Match(p.RelationshipId, @"rId(?<rId>\d+)").Groups["rId"].Value)
                            .Max(Convert.ToDecimal);

                        return $"rId{++imageId}";
                    }
                default:
                    return null;
            }
        }

        /// <summary>
        /// Calculates the maximum docPr id. The identifier is
        /// unique throughout the document. This method
        /// scans the whole document, finds and stores the max number (id is signed
        /// 23 bit integer).
        /// </summary>
        /// <param name="wordDoc">The word document.</param>
        /// <returns>System.Decimal.</returns>
        private static decimal GetMaxDocPrId(WordprocessingDocument wordDoc)
        {
            var idsList = new List<string>();
            foreach (var part in wordDoc.ContentParts())
            {
                idsList.AddRange(part.GetXDocument().Descendants(WP.docPr)
                    .SelectMany(e => e.Attributes().Where(a => a.Name == NoNamespace.id))
                    .Select(v => v.Value));
            }
            return idsList.Count == 0 ? 0 : idsList.Max(Convert.ToDecimal);
        }

        private const string InvalidImageId = "InvalidImageId";

        /// <summary>
        /// Fixes docPrIds for the document. The identifier is unique throughout the
        /// document. This method scans the whole document, finds and replaces the
        /// image ids which were marked as invalid with incremental id 
        /// (id is signed 23 bit integer).
        /// </summary>
        /// <param name="wDoc">The word processing document.</param>
        /// <param name="maxDocPrId">The current maximum document pr identifier calculated 
        /// before the document has been processed.</param>
        private static void FixUpDocPrIds(WordprocessingDocument wDoc, decimal maxDocPrId)
        {
            var elementToFind = WP.docPr;
            var docPrToChange = wDoc
                .ContentParts()
                .Select(cp => cp.GetXDocument())
                .Select(xd => xd.Descendants().Where(d => d.Name == elementToFind))
                .SelectMany(m => m);
            var nextId = maxDocPrId;
            foreach (var item in docPrToChange)
            {
                var idAtt = item.Attribute(NoNamespace.id);
                if (idAtt is { Value: InvalidImageId })
                    idAtt.Value = $"{++nextId}";
            }
            foreach (var cp in wDoc.ContentParts())
                cp.PutXDocument();
        }

        // shape type identifier
        private static int s_shapeTypeId = 1;
        private static int GetNextShapeType() => s_shapeTypeId++;

        // shape identifier
        private static int s_shapeId = 2000;
        private static string GetNextShapeId() => $"_x0000_s{s_shapeId++}";

        /// <summary>
        /// Creates and returns the image part inside the given part. The
        /// part can be either header, footer or main document part.
        /// </summary>
        /// <param name="part">The part.</param>
        /// <param name="imagePartType">Type of the image part.</param>
        /// <param name="relationshipId">The relationship identifier.</param>
        /// <returns>ImagePart.</returns>
        private static ImagePart GetImagePart(OpenXmlPart part, ImagePartType imagePartType, string relationshipId) =>
            part switch
            {
                MainDocumentPart mainDocumentPart => mainDocumentPart.AddImagePart(imagePartType, relationshipId),
                HeaderPart headerPart => headerPart.AddImagePart(imagePartType, relationshipId),
                FooterPart footerPart => footerPart.AddImagePart(imagePartType, relationshipId),
                _ => null
            };

        /// <summary>
        /// Method processes the image content and generates image element
        /// </summary>
        /// <param name="element">Source element</param>
        /// <param name="data">Data element with content</param>
        /// <param name="templateError">Error indicator</param>
        /// <param name="part">The part where the image is getting processed.</param>
        /// <returns>Image element</returns>
        private static object ProcessImageContent(XElement element, XElement data, TemplateError templateError, OpenXmlPart part)
        {
            // check for misplaced sdt content, should contain the paragraph and not vice versa
            var sdt = element.Descendants(W.sdt).FirstOrDefault();
            // get the original element with all the formatting
            var orig = sdt == null ? element.Descendants(W.p).FirstOrDefault() : sdt.Descendants(W.p).FirstOrDefault();

            // check for first run having image element in it
            if (orig == null || !orig.Descendants(W.r).FirstOrDefault().Descendants(W.drawing).Any())
            {
                return CreateContextErrorMessage(element, "Image metadata is not immediately followed by an image", templateError);
            }

            // clone the paragraph, so repeating elements won't be overridden
            var para = new XElement(orig);

            // get the xpath of of the element
            var xPath = (string)element.Attribute(PA.Select);
            // get image path
            var imagePath = EvaluateXPathToString(data, xPath, false);

            // assign unique image and paragraph ids. Image id is document property Id  (wp:docPr)
            // and relationship id is rId. Their numbering is different.
            const string imageId = InvalidImageId; // Ids will be replaced with real ones later, after transform is done
            var relationshipId = GetNextImageRelationshipId(part);

            var inline =
                para.Descendants(W.drawing)
                    .Descendants(WP.inline).FirstOrDefault();
            if (inline == null)
            {
                return CreateContextErrorMessage(element, "Image: invalid picture control", templateError);
            }

            // get aspect ratio option
            var ratioAttr = inline
                .Descendants(WP.cNvGraphicFramePr)
                .Descendants(A.graphicFrameLocks).FirstOrDefault().Attribute(NoNamespace.noChangeAspect);

            var keepSourceImageAspect = (ratioAttr == null);
            var keepOriginalImageSizeElement = inline.Descendants(Pic.cNvPicPr).FirstOrDefault();
            var keepOriginalImageSize = false;

            if (keepOriginalImageSizeElement != null)
            {
                var attr = keepOriginalImageSizeElement.Attribute("preferRelativeResize");
                if (attr != null)
                {
                    keepOriginalImageSize = attr.Value == "0";
                }
            }

            // get extent
            var extent = inline
                .Descendants(WP.extent)
                .FirstOrDefault();
            var pictureExtent = inline
                .Descendants(A.graphic)
                .Descendants(A.graphicData)
                .Descendants(Pic._pic)
                .Descendants(Pic.spPr)
                .Descendants(A.xfrm)
                .Descendants(A.ext).
                FirstOrDefault();

            if (extent == null || pictureExtent == null)
            {
                return CreateContextErrorMessage(element, "Image: missing element in picture control - extent(s)", templateError);
            }

            // get docPr
            var docPr = inline.Descendants(WP.docPr).FirstOrDefault();
            if (docPr == null)
            {
                return CreateContextErrorMessage(element, "Image: missing element in picture control - docPtr", templateError);
            }

            docPr.SetAttributeValue(NoNamespace.id, imageId);
            docPr.SetAttributeValue(NoNamespace.name, "Templated Image Content");

            var blip = inline
                    .Descendants(A.graphic)
                    .Descendants(A.graphicData)
                    .Descendants(Pic.blipFill)
                    .Descendants(A.blip)
                    .FirstOrDefault();

            if (blip is null)
                return para;

            // Add the image to main document part
            var stream = Image2Stream(imagePath, out var imagePartType, out var error);
            if (stream is not null)
            {
                var ip = GetImagePart(part, imagePartType, relationshipId);
                if (ip is null)
                {
                    error = "Failed to get image part";
                    return CreateContextErrorMessage(element, string.Concat("Image: ", error), templateError);
                }
                ip.FeedData(stream);
                stream.Close();

                // access the saved image and get the dimensions
                using var savedStream = ip.GetStream(FileMode.Open);
                using var image = System.Drawing.Image.FromStream(savedStream);
                // one inch is 914400 EMUs
                // 96dpi where dot is pixel
                var pixelInEMU = 914400 / 96;
                var width = image.Width;
                var height = image.Height;

                if (keepSourceImageAspect)
                {
                    var ratio = height / (width * 1.0);
                    if (!int.TryParse(extent.Attribute(NoNamespace.cx).Value, out width))
                    {
                        return CreateContextErrorMessage(element, "Image: Invalid image attributes",
                            templateError);
                    }
                    height = (int)(width * ratio);

                    // replace attributes
                    extent.SetAttributeValue(NoNamespace.cy, height);
                    pictureExtent.SetAttributeValue(NoNamespace.cx, width);
                    pictureExtent.SetAttributeValue(NoNamespace.cy, height);
                }

                if (keepOriginalImageSize)
                {
                    width = image.Width * pixelInEMU;
                    height = image.Height * pixelInEMU;

                    // replace attributes
                    extent.SetAttributeValue(NoNamespace.cx, width);
                    extent.SetAttributeValue(NoNamespace.cy, height);
                    pictureExtent.SetAttributeValue(NoNamespace.cx, width);
                    pictureExtent.SetAttributeValue(NoNamespace.cy, height);
                }
            }
            else
            {
                return CreateContextErrorMessage(element, string.Concat("Image: ", error), templateError);
            }

            blip.SetAttributeValue(R.embed, relationshipId);

            return para;
        }

        /// <summary>
        /// Determines whether the input image is base64 encoded string or path
        /// </summary>
        /// <param name="inputImage">Input image (either image path or base64 encoded string). Base 64 encoded string
        /// should start with MIME data type identifier followed by raw data. Example:
        /// data:image/jpg;base64,/9j/4AAQSkZJRgAB...</param>
        /// <param name="imagePartType">Image Part Type to be embedded in the document and to be
        /// referenced by image control</param>
        /// <param name="error">Error message</param>
        private static Stream Image2Stream(string inputImage, out ImagePartType imagePartType, out string error)
        {
            string imageType;
            Stream stream;
            if (inputImage.StartsWith("data:image"))
            {
                // assume the image is base64 encoded format. See https://en.wikipedia.org/wiki/Data_URI_scheme

                // get the image type and data
                imageType = Regex.Match(inputImage, @"data:image/(?<type>.+?);").Groups["type"].Value;
                var base64Data = Regex.Match(inputImage, @"data:image/(?<type>.+?),(?<data>.+)").Groups["data"].Value;

                try
                {
                    var imageBytes = Base64.ConvertFromBase64(string.Empty, base64Data);

                    stream = new MemoryStream(imageBytes, 0, imageBytes.Length);
                }
                catch (Exception)
                {
                    imagePartType = default(ImagePartType);
                    error = "Invalid Image data format";
                    return null;
                }
            }
            else
            {
                // assume this is path fo file, so get the extension
                imageType = Path.GetExtension(inputImage).Trim('.');

                try
                {
                    stream = File.Open(inputImage, FileMode.Open);
                }
                catch
                {
                    imagePartType = default;
                    error = "Invalid Image path";
                    return null;
                }
            }

            switch (imageType)
            {
                case "jpg":
                case "jpeg":
                    imagePartType = ImagePartType.Jpeg;
                    break;
                case "png":
                    imagePartType = ImagePartType.Png;
                    break;
                case "tif":
                case "tiff":
                    imagePartType = ImagePartType.Tiff;
                    break;
                case "bmp":
                    imagePartType = ImagePartType.Bmp;
                    break;
                default:
                    imagePartType = default;
                    error = "Invalid image type";
                    return null;
            }

            error = string.Empty;
            return stream;
        }

        /// <summary>
        /// Method processes internal paragraphs (marked with a prefix)
        /// </summary>
        /// <param name="element">Source element</param>
        /// <param name="data">Data element with content</param>
        /// <param name="templateError">Error indicator</param>
        /// <returns>Processed element</returns>
        private static object ProcessAParagraph(XElement element, XElement data, TemplateError templateError)
        {
            var xPath = (string)element.Attribute(PA.Select);
            var optionalString = (string)element.Attribute(PA.Optional);
            var optional = (optionalString != null && optionalString.ToLower() == "true");

            string[] newValues;
            try
            {
                newValues = EvaluateXPath(data, xPath, optional);
            }
            catch (XPathException e)
            {
                return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
            }

            var para = element.Descendants(A.r).FirstOrDefault();
            if (para is null)
                return null;

            var p = new XElement(A.r, para.Elements(A.rPr));
            var rPr = para.Elements(A.t).Elements(A.rPr).FirstOrDefault();

            var lines = newValues.SelectMany(x => x.Split('\n'));
            foreach (var line in lines)
            {
                p.Add(new XElement(A.t, rPr, line));
            }
            return p;
        }

        static object ContentReplacementTransform(XNode node, XElement data, TemplateError templateError, OpenXmlPart part)
        {
            if (node is not XElement element)
                return node;

            // TODO: need to figure out potentially better place for handling Alternate Content
            if (element.Name == MC.AlternateContent)
            {
                // assign new DrawingML object id (for repeated content)
                var docProperties = element
                    .Descendants(W.drawing)
                    .Descendants(WP.docPr)
                    .FirstOrDefault();
                docProperties?.SetAttributeValue(NoNamespace.id, InvalidImageId);

                // get the fallback picture element
                var picture = element
                    .Descendants(MC.Fallback)
                    .Descendants(W.pict)
                    .FirstOrDefault();
                if (picture is not null)
                {
                    // get the shape type element (it's okay not to have it, 
                    // as the shape might use the type defined previously and left
                    // in other shape after copy-paste operation in the editor)
                    var shapeType = picture.Descendants(VML.shapetype).FirstOrDefault();
                    var shape = picture.Descendants(VML.shape).FirstOrDefault();

                    if (shape is not null)
                    {
                        shape.SetAttributeValue(NoNamespace.id, GetNextShapeId());

                        if (shapeType is not null)
                        {
                            // get next available shape type
                            var spt = GetNextShapeType();
                            var shapeTypeId = $"_x0000_t{spt}";

                            // replace the attribute in shape type and in the corresponding shapes
                            shapeType.SetAttributeValue(O.spt, $"{spt}");
                            shapeType.SetAttributeValue(NoNamespace.id, shapeTypeId);

                            shape.SetAttributeValue(NoNamespace.type, $"#{shapeTypeId}");
                        }
                    }
                }
            }
            if (element.Name == PA.Image)
            {
                return ProcessImageContent(element, data, templateError, part);
            }
            if (element.Name == PA.Content)
            {
                if (element.Descendants(A.r).FirstOrDefault() is not null)
                {
                    return ProcessAParagraph(element, data, templateError);
                }

                var para = element.Descendants(W.p).FirstOrDefault();
                var run = element.Descendants(W.r).FirstOrDefault();

                var xPath = (string)element.Attribute(PA.Select);
                var optionalString = (string)element.Attribute(PA.Optional);
                var optional = (optionalString != null && optionalString.ToLower() == "true");

                string[] newValues;
                try
                {
                    newValues = EvaluateXPath(data, xPath, optional);
                }
                catch (XPathException e)
                {
                    return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                }

                var lines = newValues.SelectMany(x => x.Split('\n'));
                if (para is not null)
                {
                    var p = new XElement(W.p, para.Elements(W.pPr));
                    var rPr = para.Elements(W.r).Elements(W.rPr).FirstOrDefault();
                    foreach (var line in lines)
                    {
                        p.Add(new XElement(W.r, rPr,
                            (p.Elements().Count() > 1) ? new XElement(W.br) : null,
                            new XElement(W.t, line)));
                    }
                    return p;
                }
                else
                {
                    var list = new List<XElement>();
                    var rPr = run.Elements().Where(e => e.Name != W.t);
                    foreach (var line in lines)
                    {
                        list.Add(new XElement(W.r, rPr,
                            (list.Count > 0) ? new XElement(W.br) : null,
                            new XElement(W.t, line)));
                    }
                    return list;
                }
            }
            if (element.Name == PA.Repeat)
            {
                var selector = (string)element.Attribute(PA.Select);
                var optionalString = (string)element.Attribute(PA.Optional);
                var optional = (optionalString != null && optionalString.ToLower() == "true");
                var alignmentOption = (string)element.Attribute(PA.Align) ?? "vertical";

                IList<XElement> repeatingData;
                try
                {
                    repeatingData = data.XPathSelectElements(selector).ToList();
                }
                catch (XPathException e)
                {
                    return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                }
                if (!repeatingData.Any())
                {
                    if (optional)
                    {
                        return null;
                        //XElement para = element.Descendants(W.p).FirstOrDefault();
                        //if (para != null)
                        //    return new XElement(W.p, new XElement(W.r));
                        //else
                        //    return new XElement(W.r);
                    }
                    return CreateContextErrorMessage(element, "Repeat: Select returned no data", templateError);
                }
                var newContent = repeatingData.Select(d =>
                    {
                        var content = element
                            .Elements()
                            .Select(e => ContentReplacementTransform(e, d, templateError, part))
                            .ToList();
                        return content;
                    })
                    .ToList();
                switch (alignmentOption.ToLower())
                {
                    case "horizontal":
                        // keep the properties of first paragraph
                        var pPr = new XElement(W.p, newContent.First())
                            .Elements(W.p)
                            .FirstOrDefault()
                            .Elements(W.pPr)
                            .FirstOrDefault();
                        // create runs from repeated content
                        var runs = newContent.Select(x =>
                        {
                            var run = new XElement(W.p, x);
                            return run.Descendants(W.r).FirstOrDefault();
                        });
                        return pPr == null ? new XElement(W.p, runs) : new XElement(W.p, pPr, runs);
                    case "vertical":
                        return newContent;
                    default:
                        return CreateContextErrorMessage(element, "Repeat: Invalid Align option", templateError);
                }
            }
            if (element.Name == PA.Table)
            {
                IList<XElement> tableData;
                try
                {
                    tableData = data.XPathSelectElements((string)element.Attribute(PA.Select)).ToList();
                }
                catch (XPathException e)
                {
                    return CreateContextErrorMessage(element, "XPathException: " + e.Message, templateError);
                }
                if (!tableData.Any())
                    return CreateContextErrorMessage(element, "Table Select returned no data", templateError);
                var table = element.Element(W.tbl);
                var protoRow = table.Elements(W.tr).Skip(1).FirstOrDefault();
                var footerRowsBeforeTransform = table
                    .Elements(W.tr)
                    .Skip(2)
                    .ToList();
                var footerRows = footerRowsBeforeTransform
                    .Select(x => ContentReplacementTransform(x, data, templateError, part))
                    .ToList();
                if (protoRow == null)
                    return CreateContextErrorMessage(element, "Table does not contain a prototype row", templateError);
                protoRow.Descendants(W.bookmarkStart).Remove();
                protoRow.Descendants(W.bookmarkEnd).Remove();
                var newTable = new XElement(W.tbl,
                    table.Elements().Where(e => e.Name != W.tr),
                    table.Elements(W.tr).FirstOrDefault(),
                    tableData.Select(d =>
                        new XElement(W.tr,
                            protoRow.Elements().Where(r => r.Name != W.tc),
                            protoRow.Elements(W.tc)
                                .Select(tc =>
                                {
                                    var paragraph = tc.Elements(W.p).FirstOrDefault();

                                    // TODO: to check for other types (if needed, of course). Also, would be nice to refactor it, say, with
                                    // TODO: different condition, for example, with switch case which checks the type of content.
                                    if (paragraph == null)
                                    {
                                        // check if this is embedded image
                                        var image = tc.Elements(PA.Image).FirstOrDefault();
                                        if (image != null)
                                        {
                                            // has to be wrapped as table cell element, since we are re-formatting the table
                                            return new XElement(W.tc, ProcessImageContent(image, d, templateError, part));
                                        }
                                    }

                                    var cellRun = paragraph.Elements(W.r).FirstOrDefault();
                                    var xPath = paragraph.Value;
                                    string[] newValues;
                                    try
                                    {
                                        newValues = EvaluateXPath(d, xPath, false);
                                    }
                                    catch (XPathException e)
                                    {
                                        var errorCell = new XElement(W.tc,
                                            tc.Elements().Where(z => z.Name != W.p),
                                            new XElement(W.p,
                                                paragraph.Element(W.pPr),
                                                CreateRunErrorMessage(e.Message, templateError)));
                                        return errorCell;
                                    }

                                    var pPr = paragraph.Element(W.pPr);
                                    var rPr = cellRun != null ? cellRun.Element(W.rPr) : new XElement(W.rPr); //if the cell was empty there is no cellRun
                                    var newCell = new XElement(W.tc,
                                        tc.Elements().Where(z => z.Name != W.p),
                                        newValues.Select(text =>
                                            new XElement(W.p, pPr,
                                                new XElement(W.r, rPr,
                                                    new XElement(W.t, text)))));
                                    return newCell;
                                }))),
                    footerRows
                );
                return newTable;
            }
            if (element.Name == PA.Conditional)
            {
                var xPath = (string)element.Attribute(PA.Select);
                var match = (string)element.Attribute(PA.Match);
                var notMatch = (string)element.Attribute(PA.NotMatch);

                if (match == null && notMatch == null)
                    return CreateContextErrorMessage(element, "Conditional: Must specify either Match or NotMatch", templateError);
                if (match != null && notMatch != null)
                    return CreateContextErrorMessage(element, "Conditional: Cannot specify both Match and NotMatch", templateError);

                string testValue;
                try
                {
                    testValue = EvaluateXPathToString(data, xPath, false);
                }
                catch (XPathException e)
                {
                    return CreateContextErrorMessage(element, e.Message, templateError);
                }

                if ((match != null && testValue == match) || (notMatch != null && testValue != notMatch))
                {
                    var content = element.Elements().Select(e => ContentReplacementTransform(e, data, templateError, part));
                    return content;
                }
                return null;
            }
            return new XElement(element.Name,
                element.Attributes(),
                element.Nodes().Select(n => ContentReplacementTransform(n, data, templateError, part)));
        }

        private static object CreateContextErrorMessage(XElement element, string errorMessage, TemplateError templateError)
        {
            var para = element.Descendants(W.p).FirstOrDefault();
            //var run = element.Descendants(W.r).FirstOrDefault();
            var errorRun = CreateRunErrorMessage(errorMessage, templateError);
            return para != null ? new XElement(W.p, errorRun) : errorRun;
        }

        private static XElement CreateRunErrorMessage(string errorMessage, TemplateError templateError)
        {
            templateError.HasError = true;
            var errorRun = new XElement(W.r,
                new XElement(W.rPr,
                    new XElement(W.color, new XAttribute(W.val, "FF0000")),
                    new XElement(W.highlight, new XAttribute(W.val, "yellow"))),
                    new XElement(W.t, errorMessage));
            return errorRun;
        }

        private static XElement CreateParaErrorMessage(string errorMessage, TemplateError templateError)
        {
            templateError.HasError = true;
            var errorPara = new XElement(W.p,
                new XElement(W.r,
                    new XElement(W.rPr,
                        new XElement(W.color, new XAttribute(W.val, "FF0000")),
                        new XElement(W.highlight, new XAttribute(W.val, "yellow"))),
                        new XElement(W.t, errorMessage)));
            return errorPara;
        }

        private static string[] EvaluateXPath(XElement element, string xPath, bool optional)
        {
            //support some cells in the table may not have an xpath expression.
            if (string.IsNullOrWhiteSpace(xPath))
                return Array.Empty<string>();

            object xPathSelectResult;
            try
            {
                xPathSelectResult = element.XPathEvaluate(xPath);
            }
            catch (XPathException e)
            {
                throw new XPathException("XPathException: " + e.Message, e);
            }

            if (xPathSelectResult is IEnumerable enumerable and not string)
            {
                var result = enumerable.Cast<XObject>().Select(x => x switch
                {
                    XElement xElement => xElement.Value,
                    XAttribute attribute => attribute.Value,
                    _ => throw new ArgumentException($"Unknown element type: {x.GetType().Name}")
                }).ToArray();

                if (result.Length == 0 && !optional)
                    throw new XPathException($"XPath expression ({xPath}) returned no results");
                return result;
            }

            return new[] { xPathSelectResult.ToString() };
        }

        private static string EvaluateXPathToString(XElement element, string xPath, bool optional)
        {
            var selectedData = EvaluateXPath(element, xPath, true);

            return selectedData.Length switch
            {
                0 when optional => string.Empty,
                0 => throw new XPathException($"XPath expression ({xPath}) returned no results"),
                > 1 => throw new XPathException($"XPath expression ({xPath}) returned more than one node"),
                _ => selectedData.First()
            };
        }
    }
}
