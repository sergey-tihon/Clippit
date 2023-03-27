using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint
{
    internal static class PresentationBuilderTools
    {
        internal static string GetSlideTitle(XElement slide)
        {
            var titleShapes = slide
                .Element(P.cSld)
                .Element(P.spTree)
                .Descendants(P.sp)
                .Where(shape => shape
                        .Element(P.nvSpPr)
                        ?.Element(P.nvPr)
                        ?.Element(P.ph)
                        ?.Attribute(NoNamespace.type)
                        ?.Value switch
                    {
                        "title" => true,
                        "ctrTitle" => true,
                        _ => false
                    })
                .ToList();

            var paragraphText = new StringBuilder();
            foreach (var shape in titleShapes)
            {
                // Get the text in each paragraph in this shape.
                foreach (var paragraph in shape.Element(P.txBody).Descendants(A.p))
                {
                    foreach (var text in paragraph.Descendants(A.t))
                    {
                        paragraphText.Append(text.Value);
                    }
                }
            }

            return paragraphText.ToString().Trim();
        }

        internal static readonly Dictionary<XName, int> s_orderPresentation =
            new()
            {
                {P.sldMasterIdLst, 10},
                {P.notesMasterIdLst, 20},
                {P.handoutMasterIdLst, 30},
                {P.sldIdLst, 40},
                {P.sldSz, 50},
                {P.notesSz, 60},
                {P.embeddedFontLst, 70},
                {P.custShowLst, 80},
                {P.photoAlbum, 90},
                {P.custDataLst, 100},
                {P.kinsoku, 120},
                {P.defaultTextStyle, 130},
                {P.modifyVerifier, 150},
                {P.extLst, 160},
            };

        private static readonly Dictionary<XName, XName[]> s_relationshipMarkup =
            new()
            {
                {A.audioFile, new[] {R.link}},
                {A.videoFile, new[] {R.link}},
                {A.quickTimeFile, new[] {R.link}},
                {A.wavAudioFile, new[] {R.embed}},
                {A.blip, new[] {R.embed, R.link}},
                {A.hlinkClick, new[] {R.id}},
                {A.hlinkMouseOver, new[] {R.id}},
                {A.hlinkHover, new[] {R.id}},
                {A.relIds, new[] {R.cs, R.dm, R.lo, R.qs}},
                {C.chart, new[] {R.id}},
                {C.externalData, new[] {R.id}},
                {C.userShapes, new[] {R.id}},
                {Cx.chart, new[] {R.id}},
                {Cx.externalData, new[] {R.id}},
                {DGM.relIds, new[] {R.cs, R.dm, R.lo, R.qs}},
                {A14.imgLayer, new[] {R.embed}},
                {P14.media, new[] {R.embed, R.link}},
                {P.oleObj, new[] {R.id}},
                {P.externalData, new[] {R.id}},
                {P.control, new[] {R.id}},
                {P.snd, new[] {R.embed}},
                {P.sndTgt, new[] {R.embed}},
                {PAV.srcMedia, new[] {R.embed, R.link}},
                {P.contentPart, new[] {R.id}},
                {VML.fill, new[] {R.id}},
                {VML.imagedata, new[] {R.href, R.id, R.pict, O.relid}},
                {VML.stroke, new[] {R.id}},
                {WNE.toolbarData, new[] {R.id}},
                {Plegacy.textdata, new[] {XName.Get("id")}},
            };

        internal static void CopyChartObjects(ChartPart oldChart, ChartPart newChart)
        {
            foreach (var dataReference in newChart.GetXDocument().Descendants(C.externalData))
            {
                var relId = dataReference.Attribute(R.id).Value;

                if (oldChart.Parts.FirstOrDefault(p => p.RelationshipId == relId) is {} oldPartIdPair)
                {
                    switch (oldPartIdPair.OpenXmlPart)
                    {
                        case EmbeddedPackagePart oldPart:
                        {
                            var newPart = newChart.AddEmbeddedPackagePart(oldPart.ContentType);
                            using (var oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                            {
                                using var newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite);
                                oldObject.CopyTo(newObject);
                            }
                            dataReference.Attribute(R.id).Set(newChart.GetIdOfPart(newPart));
                            continue;
                        }
                        case EmbeddedObjectPart oldEmbeddedObjectPart:
                        {
                            var newPart = newChart.AddEmbeddedPackagePart(oldEmbeddedObjectPart.ContentType);
                            using (var oldObject = oldEmbeddedObjectPart.GetStream(FileMode.Open, FileAccess.Read))
                            {
                                using var newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite);
                                oldObject.CopyTo(newObject);
                            }

                            var rId = newChart.GetIdOfPart(newPart);
                            dataReference.Attribute(R.id).Set(rId);

                            // following is a hack to fix the package because the Open XML SDK does not let us create
                            // a relationship from a chart with the oleObject relationship type.

                            var pkg = newChart.OpenXmlPackage.Package;
                            var fromPart = pkg.GetParts().FirstOrDefault(p => p.Uri == newChart.Uri);
                            var rel = fromPart?.GetRelationships().FirstOrDefault(p => p.Id == rId);
                            var targetUri = rel?.TargetUri;

                            fromPart?.DeleteRelationship(rId);
                            fromPart?.CreateRelationship(targetUri, System.IO.Packaging.TargetMode.Internal,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", rId);

                            continue;
                        }
                    }
                }
                else
                {
                    //ExternalRelationship oldRelationship = oldChart.GetExternalRelationship(relId);
                    var oldRel = oldChart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel is null)
                        throw new PresentationBuilderInternalException("Internal Error 0007");
                    
                    var newRid = NewRelationshipId();
                    newChart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    dataReference.Attribute(R.id).Set(newRid);
                }
            }

            foreach (var idPartPair in oldChart.Parts)
            {
                switch (idPartPair.OpenXmlPart)
                {
                    case ThemeOverridePart oldThemeOverridePart:
                        CopyPart(oldThemeOverridePart);
                        break;
                    case ChartColorStylePart oldChartColorStylePart:
                        CopyPart(oldChartColorStylePart);
                        break;
                    case ChartStylePart oldChartColorStylePart:
                        CopyPart(oldChartColorStylePart);
                        break;
                }
            }

            void CopyPart<T>(T oldPart) where T : OpenXmlPart
            {
                var newRid = NewRelationshipId();
                var newPart = newChart.AddNewPart<T>(oldPart.ContentType, newRid);
                
                using var oldStream = oldPart.GetStream(FileMode.Open, FileAccess.Read);
                using var newStream = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite);
                oldStream.CopyTo(newStream);
            }
        }
        
        internal static void CopyExtendedChartObjects(ExtendedChartPart oldChart, ExtendedChartPart newChart)
        {
            foreach (var dataReference in newChart.GetXDocument().Descendants(Cx.externalData))
            {
                var relId = dataReference.Attribute(R.id).Value;

                if (oldChart.Parts.FirstOrDefault(p => p.RelationshipId == relId) is {} oldPartIdPair)
                {
                    switch (oldPartIdPair.OpenXmlPart)
                    {
                        case EmbeddedPackagePart oldPart:
                        {
                            var newPart = newChart.AddEmbeddedPackagePart(oldPart.ContentType);
                            using (var oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                            {
                                using var newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite);
                                oldObject.CopyTo(newObject);
                            }
                            dataReference.Attribute(R.id).Set(newChart.GetIdOfPart(newPart));
                            continue;
                        }
                        case EmbeddedObjectPart oldEmbeddedObjectPart:
                        {
                            var newPart = newChart.AddEmbeddedPackagePart(oldEmbeddedObjectPart.ContentType);
                            using (var oldObject = oldEmbeddedObjectPart.GetStream(FileMode.Open, FileAccess.Read))
                            {
                                using var newObject = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite);
                                oldObject.CopyTo(newObject);
                            }

                            var rId = newChart.GetIdOfPart(newPart);
                            dataReference.Attribute(R.id).Set(rId);

                            // following is a hack to fix the package because the Open XML SDK does not let us create
                            // a relationship from a chart with the oleObject relationship type.

                            var pkg = newChart.OpenXmlPackage.Package;
                            var fromPart = pkg.GetParts().FirstOrDefault(p => p.Uri == newChart.Uri);
                            var rel = fromPart?.GetRelationships().FirstOrDefault(p => p.Id == rId);
                            var targetUri = rel?.TargetUri;

                            fromPart?.DeleteRelationship(rId);
                            fromPart?.CreateRelationship(targetUri, System.IO.Packaging.TargetMode.Internal,
                                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject", rId);

                            continue;
                        }
                    }
                }
                else
                {
                    //ExternalRelationship oldRelationship = oldChart.GetExternalRelationship(relId);
                    var oldRel = oldChart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel is null)
                        throw new PresentationBuilderInternalException("Internal Error 0007");
                    
                    var newRid = NewRelationshipId();
                    newChart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    dataReference.Attribute(R.id).Set(newRid);
                }
            }

            foreach (var idPartPair in oldChart.Parts)
            {
                switch (idPartPair.OpenXmlPart)
                {
                    case ThemeOverridePart oldThemeOverridePart:
                        CopyPart(oldThemeOverridePart);
                        break;
                    case ChartColorStylePart oldChartColorStylePart:
                        CopyPart(oldChartColorStylePart);
                        break;
                    case ChartStylePart oldChartColorStylePart:
                        CopyPart(oldChartColorStylePart);
                        break;
                }
            }

            void CopyPart<T>(T oldPart) where T : OpenXmlPart
            {
                var newRid = NewRelationshipId();
                var newPart = newChart.AddNewPart<T>(oldPart.ContentType, newRid);
                
                using var oldStream = oldPart.GetStream(FileMode.Open, FileAccess.Read);
                using var newStream = newPart.GetStream(FileMode.Create, FileAccess.ReadWrite);
                oldStream.CopyTo(newStream);
            }
        }

        private static void UpdateContent(IEnumerable<XElement> newContent, XName elementToModify, string oldRid, string newRid)
        {
            foreach (var attributeName in s_relationshipMarkup[elementToModify])
            {
                var elementsToUpdate = newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid);
                foreach (var element in elementsToUpdate)
                    element.Attribute(attributeName).Set(newRid);
            }
        }

        private static void RemoveContent(IEnumerable<XElement> newContent, XName elementToModify, string oldRid)
        {
            foreach (var attributeName in s_relationshipMarkup[elementToModify])
            {
                newContent
                    .Descendants(elementToModify)
                    .Where(e => (string)e.Attribute(attributeName) == oldRid).Remove();
            }
        }

        internal static void AddRelationships(OpenXmlPart oldPart, OpenXmlPart newPart, IEnumerable<XElement> newContent)
        {
            var relevantElements = newContent.DescendantsAndSelf()
                .Where(d => s_relationshipMarkup.ContainsKey(d.Name) &&
                            d.Attributes().Any(a => s_relationshipMarkup[d.Name].Contains(a.Name)))
                .ToList();
            foreach (var e in relevantElements)
            {
                if (e.Name == A.hlinkClick || e.Name == A.hlinkHover || e.Name == A.hlinkMouseOver)
                {
                    var relId = (string)e.Attribute(R.id);
                    if (string.IsNullOrEmpty(relId))
                    {
                        // handle the following:
                        //<a:hlinkClick r:id=""
                        //              action="ppaction://customshow?id=0" />
                        if (e.Attribute("action") is {} attr)
                        {
                            if (attr.Value.Contains("customshow"))
                                attr.Remove();
                        }
                        continue;
                    }
                    var tempHyperlink = newPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (tempHyperlink is {})
                        continue;
                    var newRid = NewRelationshipId();
                    var oldHyperlink = oldPart.HyperlinkRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldHyperlink is null) {
                        //TODO Issue with reference to another part: var temp = oldPart.GetPartById(relId);
                        RemoveContent(newContent, e.Name, relId);
                        continue;
                    }
                    newPart.AddHyperlinkRelationship(oldHyperlink.Uri, oldHyperlink.IsExternal, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                else if (e.Name == VML.imagedata)
                {
                    var relId = (string)e.Attribute(R.href);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    if (newPart.ExternalRelationships.Any(h => h.Id == relId))
                        continue;
                    var newRid = NewRelationshipId();
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel is null)
                        throw new PresentationBuilderInternalException("Internal Error 0006");
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
                else if (e.Name == A.blip || e.Name == A14.imgLayer || e.Name == A.audioFile || e.Name == A.videoFile || e.Name == A.quickTimeFile)
                {
                    var relId = (string)e.Attribute(R.link);
                    if (string.IsNullOrEmpty(relId))
                        continue;
                    if (newPart.ExternalRelationships.Any(h => h.Id == relId))
                        continue;
                    var newRid = NewRelationshipId();
                    var oldRel = oldPart.ExternalRelationships.FirstOrDefault(h => h.Id == relId);
                    if (oldRel is null)
                        continue;
                    newPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newRid);
                    UpdateContent(newContent, e.Name, relId, newRid);
                }
            }
        }
        
        internal static void CopyRelatedMediaExternalRelationship(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement imageReference, XName attributeName)
        {
            var relId = (string)imageReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId)
                || newContentPart.ExternalRelationships.Any(er => er.Id == relId))
                return;

            var oldRel = oldContentPart.ExternalRelationships.FirstOrDefault(dpr => dpr.Id == relId);
            if (oldRel is null)
                return;

            var newId = NewRelationshipId();
            newContentPart.AddExternalRelationship(oldRel.RelationshipType, oldRel.Uri, newId);

            imageReference.Attribute(attributeName).Set(newId);
        }
        
        internal static void CopyInkPart(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement contentPartReference, XName attributeName)
        {
            var relId = (string)contentPartReference.Attribute(attributeName);
            if (newContentPart.HasRelationship(relId))
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = NewRelationshipId();
            var newPart = newContentPart.AddNewPart<CustomXmlPart>("application/inkml+xml", newId);

            using (var stream = oldPart.GetStream())
                newPart.FeedData(stream);
            contentPartReference.Attribute(attributeName).Set(newId);
        }

        internal static void CopyActiveXPart(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement activeXPartReference, XName attributeName)
        {
            var relId = (string)activeXPartReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId)
                || newContentPart.Parts.Any(p => p.RelationshipId == relId))
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = NewRelationshipId();
            var newPart = newContentPart.AddNewPart<EmbeddedControlPersistencePart>("application/vnd.ms-office.activeX+xml", newId);

            using(var stream = oldPart.GetStream())
                newPart.FeedData(stream);
            activeXPartReference.Attribute(attributeName).Set(newId);

            if (newPart.ContentType == "application/vnd.ms-office.activeX+xml")
            {
                var axc = newPart.GetXDocument();
                if (axc.Root?.Attribute(R.id) is {} attr)
                {
                    var oldPersistencePart = oldPart.GetPartById(attr.Value);

                    var newId2 = NewRelationshipId();
                    var newPersistencePart = newPart.AddNewPart<EmbeddedControlPersistenceBinaryDataPart>("application/vnd.ms-office.activeX", newId2);

                    using (var stream = oldPersistencePart.GetStream())
                        newPersistencePart.FeedData(stream);
                    attr.Set(newId2);
                }
            }
        }
        
        internal static void CopyLegacyDiagramText(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement textDataReference, XName attributeName)
        {
            var relId = (string)textDataReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId)
                || newContentPart.Parts.Any(p => p.RelationshipId == relId))
                return;

            var oldPart = oldContentPart.GetPartById(relId);

            var newId = NewRelationshipId();
            var newPart = newContentPart.AddNewPart<LegacyDiagramTextPart>(newId);

            using (var stream = oldPart.GetStream())
                newPart.FeedData(stream);
            textDataReference.Attribute(attributeName).Set(newId);
        }
        
        internal static void CopyExtendedPart(OpenXmlPart oldContentPart, OpenXmlPart newContentPart, XElement extendedReference, XName attributeName)
        {
            var relId = (string)extendedReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId))
                return;
            try
            {
                // First look to see if this relId has already been added to the new document.
                // This is necessary for those parts that get processed with both old and new ids, such as the comments
                // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
                // in that case.
                if (newContentPart.HasRelationship(relId))
                    return;

                var oldPart = (ExtendedPart)oldContentPart.GetPartById(relId);
                var fileInfo = new FileInfo(oldPart.Uri.OriginalString);

                var newPart = newContentPart switch
                {
                    ChartColorStylePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ChartDrawingPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ChartPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ChartsheetPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ChartStylePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CommentAuthorsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ConnectionsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ControlPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CoreFilePropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomDataPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomFilePropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomizationPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomPropertyPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomUIPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomXmlMappingsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomXmlPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    CustomXmlPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramColorsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramLayoutDefinitionPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramPersistLayoutPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DiagramStylePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DigitalSignatureOriginPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    DrawingsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    EmbeddedControlPersistenceBinaryDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    EmbeddedControlPersistencePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    EmbeddedObjectPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    EmbeddedPackagePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ExtendedFilePropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ExtendedPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    FontPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    FontTablePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    HandoutMasterPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    InternationalMacroSheetPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    LegacyDiagramTextInfoPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    LegacyDiagramTextPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    MacroSheetPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    NotesMasterPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    NotesSlidePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    PresentationPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    PresentationPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    RibbonAndBackstageCustomizationsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SingleCellTablePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlideCommentsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlideLayoutPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlideMasterPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlidePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    SlideSyncDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    StyleDefinitionsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType,fileInfo.Extension),
                    StylesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    TableDefinitionPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    TableStylesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ThemeOverridePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ThemePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ThumbnailPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    TimeLineCachePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    TimeLinePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    UserDefinedTagsPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    VbaDataPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    VbaProjectPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    ViewPropertiesPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    VmlDrawingPart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    XmlSignaturePart part => part.AddExtendedPart(oldPart.RelationshipType, oldPart.ContentType, fileInfo.Extension),
                    _ => null
                };

                relId = newContentPart.GetIdOfPart(newPart);
                using (var stream = oldPart.GetStream())
                    newPart?.FeedData(stream);
                extendedReference.Attribute(attributeName).Set(relId);
            }
            catch (ArgumentOutOfRangeException)
            {
                try
                {
                    var er = oldContentPart.GetExternalRelationship(relId);
                    var newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                    extendedReference.Attribute(R.id).Value = newEr.Id;
                }
                catch (KeyNotFoundException)
                {
                    var newPart = newContentPart.OpenXmlPackage.Package.GetParts().FirstOrDefault(p => p.Uri == newContentPart.Uri);
                    if (newPart.RelationshipExists(relId) == false)
                    {
                        newPart.CreateRelationship(new Uri("NULL", UriKind.RelativeOrAbsolute),
                            System.IO.Packaging.TargetMode.Internal,
                            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", relId);
                    }
                }
            }
        }
        
        internal static void CopyRelatedSound(PresentationDocument newDocument, OpenXmlPart oldContentPart, OpenXmlPart newContentPart,
            XElement soundReference, XName attributeName)
        {
            var relId = (string)soundReference.Attribute(attributeName);
            if (string.IsNullOrEmpty(relId)
                || newContentPart.ExternalRelationships.Any(er => er.Id == relId)
                || newContentPart.DataPartReferenceRelationships.Any(er => er.Id == relId))
                return;

            switch (oldContentPart.GetReferenceRelationship(relId))
            {
                case AudioReferenceRelationship audioRef:
                {
                    var newSound = newDocument.CreateMediaDataPart(audioRef.DataPart.ContentType);
                    using (var stream = audioRef.DataPart.GetStream())
                        newSound.FeedData(stream);

                    var newRel = newContentPart switch
                    {
                        SlidePart part => part.AddAudioReferenceRelationship(newSound),
                        SlideLayoutPart part => part.AddAudioReferenceRelationship(newSound),
                        SlideMasterPart part => part.AddAudioReferenceRelationship(newSound),
                        HandoutMasterPart part => part.AddAudioReferenceRelationship(newSound),
                        NotesMasterPart part => part.AddAudioReferenceRelationship(newSound),
                        NotesSlidePart part => part.AddAudioReferenceRelationship(newSound),
                        _ => null
                    };
                    soundReference.Attribute(attributeName).Set(newRel?.Id);
                    break;
                }
                case MediaReferenceRelationship mediaRef:
                {
                    var newSound = newDocument.CreateMediaDataPart(mediaRef.DataPart.ContentType);
                    using (var stream = mediaRef.DataPart.GetStream())
                        newSound.FeedData(stream);

                    var newRel = newContentPart switch
                    {
                        SlidePart part => part.AddMediaReferenceRelationship(newSound),
                        SlideLayoutPart part => part.AddMediaReferenceRelationship(newSound),
                        SlideMasterPart part => part.AddMediaReferenceRelationship(newSound),
                        _ => null
                    };
                    soundReference.Attribute(attributeName).Set(newRel?.Id);
                    break;
                }
            }
        }
        
        internal static void Set(this XAttribute attr, string value)
        {
            if (attr is null) return;
            attr.Value = value;
        }
        
        internal static bool HasRelationship(this OpenXmlPart part, string relId) =>
            string.IsNullOrEmpty(relId)
            || part.Parts.Any(p => p.RelationshipId == relId)
            || part.ExternalRelationships.Any(er => er.Id == relId);

        internal static string NewRelationshipId() =>
            "rcId" + Guid.NewGuid().ToString().Replace("-", "").Substring(0, 16);
    }
    
    public class PresentationBuilderException : Exception
    {
        public PresentationBuilderException(string message) : base(message) { }
    }

    public class PresentationBuilderInternalException : Exception
    {
        public PresentationBuilderInternalException(string message) : base(message) { }
    }
}
