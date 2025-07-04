﻿using System.Xml;
using System.Xml.Linq;
using Clippit.Internal;
using DocumentFormat.OpenXml.Experimental;
using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Packaging;
using Path = System.IO.Path;
using PBT = Clippit.PowerPoint.Fluent.PresentationBuilderTools;

namespace Clippit.PowerPoint.Fluent;

internal sealed partial class FluentPresentationBuilder
{
    private void CopyStartingParts(PresentationDocument sourceDocument)
    {
        // A Core File Properties part does not have implicit or explicit relationships to other parts.
        var srcCorePart = sourceDocument.CoreFilePropertiesPart;
        if (srcCorePart?.GetXDocument().Root is not null && _newDocument.CoreFilePropertiesPart is null)
        {
            _newDocument.AddCoreFilePropertiesPart();
            var newXDoc = _newDocument.CoreFilePropertiesPart.GetXDocument();
            newXDoc.Declaration.Standalone = "yes";
            newXDoc.Declaration.Encoding = "UTF-8";
            var sourceXDoc = srcCorePart.GetXDocument();
            newXDoc.Add(sourceXDoc.Root);
        }

        // An application attributes part does not have implicit or explicit relationships to other parts.
        if (sourceDocument.ExtendedFilePropertiesPart is { } extPart && _newDocument.ExtendedFilePropertiesPart is null)
        {
            _newDocument.AddExtendedFilePropertiesPart();
            var newXDoc = _newDocument.ExtendedFilePropertiesPart.GetXDocument();
            newXDoc.Declaration.Standalone = "yes";
            newXDoc.Declaration.Encoding = "UTF-8";
            newXDoc.Add(extPart.GetXDocument().Root);
        }

        // An custom file properties part does not have implicit or explicit relationships to other parts.
        if (sourceDocument.CustomFilePropertiesPart is { } customPart && _newDocument.CustomFilePropertiesPart is null)
        {
            _newDocument.AddCustomFilePropertiesPart();
            var newXDoc = _newDocument.CustomFilePropertiesPart.GetXDocument();
            newXDoc.Declaration.Standalone = "yes";
            newXDoc.Declaration.Encoding = "UTF-8";
            newXDoc.Add(customPart.GetXDocument().Root);
        }
    }

#if false
            // TODO need to handle the following

            { P.custShowLst, 80 },
            { P.photoAlbum, 90 },
            { P.custDataLst, 100 },
            { P.kinsoku, 120 },
            { P.modifyVerifier, 150 },
#endif
    // Copy handout master, notes master, presentation properties and view properties, if they exist
    private void CopyPresentationParts(PresentationDocument sourceDocument)
    {
        var newPresentation = _newDocument.PresentationPart.GetXDocument();
        var newPresentationRoot = newPresentation.Root!;

        // Copy slide and note slide sizes
        var oldPresentationDoc = sourceDocument.PresentationPart.GetXDocument();

        foreach (var att in oldPresentationDoc.Root.Attributes())
        {
            if (!att.IsNamespaceDeclaration && newPresentationRoot.Attribute(att.Name) is null)
                newPresentationRoot.Add(oldPresentationDoc.Root.Attribute(att.Name));
        }

        if (
            oldPresentationDoc.Root.Elements(P.sldSz).FirstOrDefault() is { } oldElement
            && newPresentationRoot.Element(P.sldSz) is null
        )
        {
            newPresentationRoot.Add(oldElement);
        }

        // Copy Font Parts
        if (
            oldPresentationDoc.Root.Element(P.embeddedFontLst) is { } embeddedFontLst
            && newPresentationRoot.Element(P.embeddedFontLst) is null
        )
        {
            var newFontLst = new XElement(P.embeddedFontLst);
            foreach (var font in embeddedFontLst.Elements(P.embeddedFont))
            {
                var newEmbeddedFont = new XElement(P.embeddedFont, font.Elements(P.font));

                if (font.Element(P.regular) is not null)
                    newEmbeddedFont.Add(CreateEmbeddedFontPart(sourceDocument, font, P.regular));
                if (font.Element(P.bold) is not null)
                    newEmbeddedFont.Add(CreateEmbeddedFontPart(sourceDocument, font, P.bold));
                if (font.Element(P.italic) is not null)
                    newEmbeddedFont.Add(CreateEmbeddedFontPart(sourceDocument, font, P.italic));
                if (font.Element(P.boldItalic) is not null)
                    newEmbeddedFont.Add(CreateEmbeddedFontPart(sourceDocument, font, P.boldItalic));

                newFontLst.Add(newEmbeddedFont);
            }
            newPresentationRoot.Add(newFontLst);
        }

        if (newPresentationRoot.Element(P.defaultTextStyle) is null)
            newPresentationRoot.Add(oldPresentationDoc.Root.Element(P.defaultTextStyle));
        if (newPresentationRoot.Element(P.extLst) is null)
            newPresentationRoot.Add(SanitizeExtLst(oldPresentationDoc.Root.Elements(P.extLst)));

        //<p:embeddedFont xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
        //                         xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        //  <p:font typeface="Perpetua" panose="02020502060401020303" pitchFamily="18" charset="0" />
        //  <p:regular r:id="rId5" />
        //  <p:bold r:id="rId6" />
        //  <p:italic r:id="rId7" />
        //  <p:boldItalic r:id="rId8" />
        //</p:embeddedFont>

        // Copy Handout Master
        if (
            sourceDocument.PresentationPart.HandoutMasterPart is { } oldMaster
            && _newDocument.PresentationPart.HandoutMasterPart is null
        )
        {
            var newMaster = _newDocument.PresentationPart.AddNewPart<HandoutMasterPart>();

            // Copy theme for master
            var newThemePart = newMaster.AddNewPart<ThemePart>();
            newThemePart.PutXDocument(new XDocument(oldMaster.ThemePart.GetXDocument()));
            PBT.AddRelationships(oldMaster.ThemePart, newThemePart, [newThemePart.GetXDocument().Root]);
            CopyRelatedPartsForContentParts(oldMaster.ThemePart, newThemePart, [newThemePart.GetXDocument().Root]);

            // Copy master
            newMaster.PutXDocument(new XDocument(oldMaster.GetXDocument()));
            PBT.AddRelationships(oldMaster, newMaster, [newMaster.GetXDocument().Root]);
            CopyRelatedPartsForContentParts(oldMaster, newMaster, [newMaster.GetXDocument().Root]);

            newPresentationRoot.Add(
                new XElement(
                    P.handoutMasterIdLst,
                    new XElement(
                        P.handoutMasterId,
                        new XAttribute(R.id, _newDocument.PresentationPart.GetIdOfPart(newMaster))
                    )
                )
            );
        }

        // Copy Notes Master
        CopyNotesMaster(sourceDocument);

        // Copy Presentation Properties
        if (
            sourceDocument.PresentationPart.PresentationPropertiesPart is { } presentationPropertiesPart
            && _newDocument.PresentationPart.PresentationPropertiesPart is null
        )
        {
            var newPart = _newDocument.PresentationPart.AddNewPart<PresentationPropertiesPart>();
            var xd1 = presentationPropertiesPart.GetXDocument();
            xd1.Descendants(P.custShow).Remove();
            newPart.PutXDocument(xd1);
        }

        // Copy View Properties
        if (
            sourceDocument.PresentationPart.ViewPropertiesPart is { } viewPropertiesPart
            && _newDocument.PresentationPart.ViewPropertiesPart is null
        )
        {
            var newPart = _newDocument.PresentationPart.AddNewPart<ViewPropertiesPart>();
            var xd = viewPropertiesPart.GetXDocument();
            xd.Descendants(P.outlineViewPr).Elements(P.sldLst).Remove();
            newPart.PutXDocument(xd);
        }

        foreach (
            var legacyDocTextInfo in sourceDocument.PresentationPart.Parts.Where(p =>
                p.OpenXmlPart.RelationshipType
                == "http://schemas.microsoft.com/office/2006/relationships/legacyDocTextInfo"
            )
        )
        {
            var newPart = _newDocument.PresentationPart.AddNewPart<LegacyDiagramTextInfoPart>();
            using var stream = legacyDocTextInfo.OpenXmlPart.GetStream();
            newPart.FeedData(stream);
        }

        var listOfRootChildren = newPresentationRoot.Elements().ToList();
        foreach (var rc in listOfRootChildren)
            rc.Remove();
        newPresentationRoot.Add(listOfRootChildren.OrderBy(e => PBT.OrderPresentation.GetValueOrDefault(e.Name, 999)));
    }

    /// <summary>
    /// This method remove custom p:ext elements from the extLst element, especially ones that reference custom metadata
    /// Example:
    ///     <p:extLst>
    ///       <p:ext uri="http://customooxmlschemas.google.com/">
    ///         <go:slidesCustomData xmlns="" ... r:id="rId16" ... />
    /// </summary>
    /// <param name="extLsts">List of all <p:extLst> from source presentation.xml</param>
    /// <returns>Modified copy of all elements</returns>
    private static IEnumerable<XElement> SanitizeExtLst(IEnumerable<XElement> extLstList)
    {
        foreach (var srcExtLst in extLstList)
        {
            // Deep clone the element
            var extLst = new XElement(srcExtLst);

            // Sanitize all p:ext elements with r:Id attributes on any child element
            foreach (var ext in extLst.Elements(P.ext).ToList())
            {
                var hasRid = ext.Descendants().Any(e => e.Attributes().Any(a => a.Name == R.id));
                if (hasRid)
                    ext.Remove();
            }

            yield return extLst;
        }
    }

    private XElement CreateEmbeddedFontPart(PresentationDocument sourceDocument, XElement font, XName fontXName)
    {
        var oldFontPartId = (string)font.Element(fontXName).Attributes(R.id).FirstOrDefault();
        if (!sourceDocument.PresentationPart.TryGetPartById(oldFontPartId, out var oldFontPart))
            return null;
        if (oldFontPart is not FontPart)
            throw new FormatException($"Part {oldFontPartId} is not {nameof(FontPart)}");

        var fontPartType = oldFontPart.ContentType switch
        {
            "application/x-fontdata" => FontPartType.FontData,
            "application/x-font-ttf" => FontPartType.FontTtf,
            _ => FontPartType.FontOdttf,
        };

        var newFontPartId = Relationships.GetNewRelationshipId();
        var newFontPart = _newDocument.PresentationPart.AddFontPart(fontPartType, newFontPartId);
        using (var stream = oldFontPart.GetStream())
            newFontPart.FeedData(stream);
        return new XElement(fontXName, new XAttribute(R.id, newFontPartId));
    }

    // Copies notes master and notesSz element from presentation
    private void CopyNotesMaster(PresentationDocument sourceDocument)
    {
        // Copy notesSz element from presentation
        var newPresentation = _newDocument.PresentationPart.GetXDocument();
        var oldPresentationDoc = sourceDocument.PresentationPart.GetXDocument();
        var oldElement = oldPresentationDoc.Root.Element(P.notesSz);
        newPresentation.Root.Element(P.notesSz).ReplaceWith(oldElement);

        // Copy Notes Master
        if (
            sourceDocument.PresentationPart.NotesMasterPart is { } oldMaster
            && _newDocument.PresentationPart.NotesMasterPart is null
        )
        {
            var newMaster = _newDocument.PresentationPart.AddNewPart<NotesMasterPart>();

            // Copy theme for master
            if (oldMaster.ThemePart is { } themePart)
            {
                var newThemePart = newMaster.AddNewPart<ThemePart>();
                newThemePart.PutXDocument(new XDocument(themePart.GetXDocument()));
                PBT.AddRelationships(themePart, newThemePart, [newThemePart.GetXDocument().Root]);
                CopyRelatedPartsForContentParts(themePart, newThemePart, [newThemePart.GetXDocument().Root]);
            }

            // Copy master
            newMaster.PutXDocument(new XDocument(oldMaster.GetXDocument()));
            PBT.AddRelationships(oldMaster, newMaster, [newMaster.GetXDocument().Root]);
            CopyRelatedPartsForContentParts(oldMaster, newMaster, [newMaster.GetXDocument().Root]);

            newPresentation.Root.Add(
                new XElement(
                    P.notesMasterIdLst,
                    new XElement(
                        P.notesMasterId,
                        new XAttribute(R.id, _newDocument.PresentationPart.GetIdOfPart(newMaster))
                    )
                )
            );
        }
    }

    private void CopyComments(PresentationDocument oldDocument, SlidePart oldSlide, SlidePart newSlide)
    {
        newSlide.AddNewPart<SlideCommentsPart>();
        newSlide.SlideCommentsPart.PutXDocument(new XDocument(oldSlide.SlideCommentsPart.GetXDocument()));
        var newSlideComments = newSlide.SlideCommentsPart.GetXDocument();
        var oldAuthors = oldDocument.PresentationPart.CommentAuthorsPart.GetXDocument();
        foreach (var comment in newSlideComments.Root.Elements(P.cm))
        {
            var newAuthor = FindCommentsAuthor(comment, oldAuthors);
            // Update last index value for new comment
            comment.Attribute(NoNamespace.authorId).SetValue(newAuthor.Attribute(NoNamespace.id).Value);
            var lastIndex = Convert.ToUInt32(newAuthor.Attribute(NoNamespace.lastIdx).Value);
            comment.Attribute(NoNamespace.idx).SetValue(lastIndex.ToString());
            newAuthor.Attribute(NoNamespace.lastIdx).SetValue(Convert.ToString(lastIndex + 1));
        }
    }

    private XElement FindCommentsAuthor(XElement comment, XDocument oldAuthors)
    {
        var oldAuthor = oldAuthors
            .Root.Elements(P.cmAuthor)
            .FirstOrDefault(f => f.Attribute(NoNamespace.id).Value == comment.Attribute(NoNamespace.authorId).Value);
        if (_newDocument.PresentationPart.CommentAuthorsPart is null)
        {
            _newDocument.PresentationPart.AddNewPart<CommentAuthorsPart>();
            _newDocument.PresentationPart.CommentAuthorsPart.PutXDocument(
                new XDocument(
                    new XElement(
                        P.cmAuthorLst,
                        new XAttribute(XNamespace.Xmlns + "a", A.a),
                        new XAttribute(XNamespace.Xmlns + "r", R.r),
                        new XAttribute(XNamespace.Xmlns + "p", P.p)
                    )
                )
            );
        }
        var authors = _newDocument.PresentationPart.CommentAuthorsPart.GetXDocument();
        var newAuthor = authors
            .Root.Elements(P.cmAuthor)
            .FirstOrDefault(f =>
                f.Attribute(NoNamespace.initials).Value == oldAuthor.Attribute(NoNamespace.initials).Value
            );
        if (newAuthor is null)
        {
            uint newId = 0;
            var ids = authors.Root.Descendants(P.cmAuthor).Select(f => (uint)f.Attribute(NoNamespace.id)).ToList();
            if (ids.Count != 0)
                newId = ids.Max() + 1;

            newAuthor = new XElement(
                P.cmAuthor,
                new XAttribute(NoNamespace.id, newId.ToString()),
                new XAttribute(NoNamespace.name, oldAuthor.Attribute(NoNamespace.name).Value),
                new XAttribute(NoNamespace.initials, oldAuthor.Attribute(NoNamespace.initials).Value),
                new XAttribute(NoNamespace.lastIdx, "1"),
                new XAttribute(NoNamespace.clrIdx, newId.ToString())
            );
            authors.Root.Add(newAuthor);
        }

        return newAuthor;
    }

    private void CopyTableStyles(PresentationDocument oldDocument, OpenXmlPart newContentPart)
    {
        if (oldDocument.PresentationPart.TableStylesPart is null)
            return;

        var oldTableStylesDocument = oldDocument.PresentationPart.TableStylesPart.GetXDocument();
        var oldTableStyles = oldTableStylesDocument.Root.Elements(A.tblStyle).ToList();

        foreach (var table in newContentPart.GetXDocument().Descendants(A.tableStyleId))
        {
            var styleId = table.Value;
            if (string.IsNullOrEmpty(styleId))
                continue;

            // Find old style
            var oldStyle = oldTableStyles.FirstOrDefault(f => f.Attribute(NoNamespace.styleId).Value == styleId);
            if (oldStyle is null)
                continue;

            // Create new TableStylesPart, if needed
            XDocument tableStyles;
            if (_newDocument.PresentationPart.TableStylesPart is null)
            {
                var newStylesPart = _newDocument.PresentationPart.AddNewPart<TableStylesPart>();
                tableStyles = new XDocument(
                    new XElement(
                        A.tblStyleLst,
                        new XAttribute(XNamespace.Xmlns + "a", A.a),
                        new XAttribute(NoNamespace.def, styleId)
                    )
                );
                newStylesPart.PutXDocument(tableStyles);
            }
            else
                tableStyles = _newDocument.PresentationPart.TableStylesPart.GetXDocument();

            // Search new TableStylesPart to see if it contains the ID
            if (
                tableStyles
                    .Root.Elements(A.tblStyle)
                    .FirstOrDefault(f => f.Attribute(NoNamespace.styleId).Value == styleId)
                is not null
            )
                continue;

            // Copy style to new part
            tableStyles.Root.Add(oldStyle);
        }
    }

    private void CopyRelatedPartsForContentParts(
        OpenXmlPart oldContentPart,
        OpenXmlPart newContentPart,
        IEnumerable<XElement> newContent
    )
    {
        var relevantElements = newContent
            .DescendantsAndSelf()
            .Where(d =>
                d.Name == VML.imagedata
                || d.Name == VML.fill
                || d.Name == VML.stroke
                || d.Name == A.blip
                || d.Name == SVG.svgBlip
            )
            .ToList();
        foreach (var imageReference in relevantElements)
        {
            CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.embed);
            CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.pict);
            CopyRelatedImage(oldContentPart, newContentPart, imageReference, R.id);
            CopyRelatedImage(oldContentPart, newContentPart, imageReference, O.relid);
        }

        relevantElements = newContent
            .DescendantsAndSelf()
            .Where(d => d.Name == A.videoFile || d.Name == A.quickTimeFile)
            .ToList();
        foreach (var imageReference in relevantElements)
        {
            CopyRelatedMedia(oldContentPart, newContentPart, imageReference, R.link, "video");
        }

        relevantElements = newContent
            .DescendantsAndSelf()
            .Where(d => d.Name == P14.media || d.Name == PAV.srcMedia)
            .ToList();
        foreach (var imageReference in relevantElements)
        {
            CopyRelatedMedia(oldContentPart, newContentPart, imageReference, R.embed, "media");
            PBT.CopyRelatedMediaExternalRelationship(oldContentPart, newContentPart, imageReference, R.link);
        }

        foreach (var extendedReference in newContent.DescendantsAndSelf(A14.imgLayer))
        {
            PBT.CopyExtendedPart(oldContentPart, newContentPart, extendedReference, R.embed);
        }

        foreach (var contentPartReference in newContent.DescendantsAndSelf(P.contentPart))
        {
            PBT.CopyInkPart(oldContentPart, newContentPart, contentPartReference, R.id);
        }

        foreach (var contentPartReference in newContent.DescendantsAndSelf(P.control))
        {
            PBT.CopyActiveXPart(oldContentPart, newContentPart, contentPartReference, R.id);
        }

        foreach (var contentPartReference in newContent.DescendantsAndSelf(Plegacy.textdata))
        {
            PBT.CopyLegacyDiagramText(oldContentPart, newContentPart, contentPartReference, "id");
        }

        foreach (
            var diagramReference in newContent
                .DescendantsAndSelf()
                .Where(d => d.Name == DGM.relIds || d.Name == A.relIds)
        )
        {
            // dm attribute
            var relId = diagramReference.Attribute(R.dm).Value;
            if (newContentPart.HasRelationship(relId))
                continue;

            var oldPart = oldContentPart.GetPartById(relId);
            OpenXmlPart newPart = newContentPart.AddNewPart<DiagramDataPart>();
            newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
            diagramReference.Attribute(R.dm).Value = newContentPart.GetIdOfPart(newPart);
            PBT.AddRelationships(oldPart, newPart, [newPart.GetXDocument().Root]);
            CopyRelatedPartsForContentParts(oldPart, newPart, [newPart.GetXDocument().Root]);

            // lo attribute
            relId = diagramReference.Attribute(R.lo).Value;
            if (newContentPart.HasRelationship(relId))
                continue;

            oldPart = oldContentPart.GetPartById(relId);
            newPart = newContentPart.AddNewPart<DiagramLayoutDefinitionPart>();
            newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
            diagramReference.Attribute(R.lo).Value = newContentPart.GetIdOfPart(newPart);
            PBT.AddRelationships(oldPart, newPart, [newPart.GetXDocument().Root]);
            CopyRelatedPartsForContentParts(oldPart, newPart, [newPart.GetXDocument().Root]);

            // qs attribute
            relId = diagramReference.Attribute(R.qs).Value;
            if (newContentPart.HasRelationship(relId))
                continue;

            oldPart = oldContentPart.GetPartById(relId);
            newPart = newContentPart.AddNewPart<DiagramStylePart>();
            newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
            diagramReference.Attribute(R.qs).Value = newContentPart.GetIdOfPart(newPart);
            PBT.AddRelationships(oldPart, newPart, [newPart.GetXDocument().Root]);
            CopyRelatedPartsForContentParts(oldPart, newPart, [newPart.GetXDocument().Root]);

            // cs attribute
            relId = diagramReference.Attribute(R.cs).Value;
            if (newContentPart.HasRelationship(relId))
                continue;

            oldPart = oldContentPart.GetPartById(relId);
            newPart = newContentPart.AddNewPart<DiagramColorsPart>();
            newPart.GetXDocument().Add(oldPart.GetXDocument().Root);
            diagramReference.Attribute(R.cs).Value = newContentPart.GetIdOfPart(newPart);
            PBT.AddRelationships(oldPart, newPart, [newPart.GetXDocument().Root]);
            CopyRelatedPartsForContentParts(oldPart, newPart, [newPart.GetXDocument().Root]);
        }

        foreach (
            var oleReference in newContent
                .DescendantsAndSelf()
                .Where(d => d.Name == P.oleObj || d.Name == P.externalData)
        )
        {
            var relId = oleReference.Attribute(R.id).Value;

            // First look to see if this relId has already been added to the new document.
            // This is necessary for those parts that get processed with both old and new ids, such as the comments
            // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
            // in that case.
            if (newContentPart.HasRelationship(relId))
                continue;

            var oldPartIdPair = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (oldPartIdPair != default)
            {
                var oldPart = oldPartIdPair.OpenXmlPart;
                OpenXmlPart newPart = null;
                newPart = oldPart switch
                {
                    EmbeddedObjectPart => newContentPart switch
                    {
                        DialogsheetPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                        HandoutMasterPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                        NotesMasterPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                        NotesSlidePart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                        SlideLayoutPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                        SlideMasterPart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                        SlidePart part => part.AddEmbeddedObjectPart(oldPart.ContentType),
                        _ => newPart,
                    },
                    EmbeddedPackagePart => newContentPart switch
                    {
                        ChartPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                        HandoutMasterPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                        NotesMasterPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                        NotesSlidePart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                        SlideLayoutPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                        SlideMasterPart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                        SlidePart part => part.AddEmbeddedPackagePart(oldPart.ContentType),
                        _ => newPart,
                    },
                    _ => newPart,
                };
                using (var oldObject = oldPart.GetStream(FileMode.Open, FileAccess.Read))
                {
                    newPart.FeedData(oldObject);
                }
                oleReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
            }
            else
            {
                var er = oldContentPart.GetExternalRelationship(relId);
                var newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                oleReference.Attribute(R.id).Set(newEr.Id);
            }
        }

        foreach (var chartReference in newContent.DescendantsAndSelf(C.chart))
        {
            var relId = (string)chartReference.Attribute(R.id);
            if (newContentPart.HasRelationship(relId))
                continue;

            var oldPartIdPair2 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (oldPartIdPair2.OpenXmlPart is ChartPart oldPart)
            {
                var oldChart = oldPart.GetXDocument();
                var newPart = newContentPart.AddNewPart<ChartPart>();
                var newChart = newPart.GetXDocument();
                newChart.Add(oldChart.Root);
                chartReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                PBT.CopyChartObjects(oldPart, newPart);
                CopyRelatedPartsForContentParts(oldPart, newPart, [newChart.Root]);
            }
        }

        foreach (var chartReference in newContent.DescendantsAndSelf(Cx.chart))
        {
            var relId = (string)chartReference.Attribute(R.id);
            if (newContentPart.HasRelationship(relId))
                continue;

            var oldPartIdPair2 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (oldPartIdPair2.OpenXmlPart is ExtendedChartPart oldPart)
            {
                var oldChart = oldPart.GetXDocument();
                var newPart = newContentPart.AddNewPart<ExtendedChartPart>();
                var newChart = newPart.GetXDocument();
                newChart.Add(oldChart.Root);
                chartReference.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                PBT.CopyExtendedChartObjects(oldPart, newPart);
                CopyRelatedPartsForContentParts(oldPart, newPart, [newChart.Root]);
            }
        }

        foreach (var userShape in newContent.DescendantsAndSelf(C.userShapes))
        {
            var relId = (string)userShape.Attribute(R.id);
            if (newContentPart.HasRelationship(relId))
                continue;

            var oldPartIdPair3 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (oldPartIdPair3.OpenXmlPart is ChartDrawingPart oldPart)
            {
                var oldXDoc = oldPart.GetXDocument();
                var newPart = newContentPart.AddNewPart<ChartDrawingPart>();
                var newXDoc = newPart.GetXDocument();
                newXDoc.Add(oldXDoc.Root);
                userShape.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
                PBT.AddRelationships(oldPart, newPart, newContent);
                CopyRelatedPartsForContentParts(oldPart, newPart, [newXDoc.Root]);
            }
        }

        foreach (var tags in newContent.DescendantsAndSelf(P.tags))
        {
            var relId = (string)tags.Attribute(R.id);
            if (newContentPart.HasRelationship(relId))
                continue;

            var oldPartIdPair4 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (oldPartIdPair4.OpenXmlPart is UserDefinedTagsPart oldPart)
            {
                var oldXDoc = oldPart.GetXDocument();
                var newPart = newContentPart.AddNewPart<UserDefinedTagsPart>();
                var newXDoc = newPart.GetXDocument();
                newXDoc.Add(oldXDoc.Root);
                tags.Attribute(R.id).Value = newContentPart.GetIdOfPart(newPart);
            }
        }

        foreach (var custData in newContent.DescendantsAndSelf(P.custData))
        {
            var relId = (string)custData.Attribute(R.id);
            if (string.IsNullOrEmpty(relId) || newContentPart.Parts.Any(p => p.RelationshipId == relId))
                continue;

            var oldPartIdPair9 = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
            if (oldPartIdPair9 != default)
            {
                var newPart = _newDocument.PresentationPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using (var stream = oldPartIdPair9.OpenXmlPart.GetStream())
                    newPart.FeedData(stream);
                foreach (
                    var itemProps in oldPartIdPair9.OpenXmlPart.Parts.Where(p =>
                        p.OpenXmlPart.ContentType
                        == "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
                    )
                )
                {
                    var newId2 = Relationships.GetNewRelationshipId();
                    var cxpp = newPart.AddNewPart<CustomXmlPropertiesPart>(
                        "application/vnd.openxmlformats-officedocument.customXmlProperties+xml",
                        newId2
                    );
                    using var stream = itemProps.OpenXmlPart.GetStream();
                    cxpp.FeedData(stream);
                }
                var newId = Relationships.GetNewRelationshipId();
                newContentPart.CreateRelationshipToPart(newPart, newId);
                custData.Attribute(R.id).Value = newId;
            }
        }

        foreach (var soundReference in newContent.DescendantsAndSelf().Where(d => d.Name == A.audioFile))
            PBT.CopyRelatedSound(_newDocument, oldContentPart, newContentPart, soundReference, R.link);

        if (
            (oldContentPart is ChartsheetPart && newContentPart is ChartsheetPart)
            || (oldContentPart is DialogsheetPart && newContentPart is DialogsheetPart)
            || (oldContentPart is HandoutMasterPart && newContentPart is HandoutMasterPart)
            || (oldContentPart is InternationalMacroSheetPart && newContentPart is InternationalMacroSheetPart)
            || (oldContentPart is MacroSheetPart && newContentPart is MacroSheetPart)
            || (oldContentPart is NotesMasterPart && newContentPart is NotesMasterPart)
            || (oldContentPart is NotesSlidePart && newContentPart is NotesSlidePart)
            || (oldContentPart is SlideLayoutPart && newContentPart is SlideLayoutPart)
            || (oldContentPart is SlideMasterPart && newContentPart is SlideMasterPart)
            || (oldContentPart is SlidePart && newContentPart is SlidePart)
            || (oldContentPart is WorksheetPart && newContentPart is WorksheetPart)
        )
        {
            foreach (
                var soundReference in newContent
                    .DescendantsAndSelf()
                    .Where(d =>
                        d.Name == P.snd
                        || d.Name == P.sndTgt
                        || d.Name == A.wavAudioFile
                        || d.Name == A.snd
                        || d.Name == PAV.srcMedia
                    )
            )
                PBT.CopyRelatedSound(_newDocument, oldContentPart, newContentPart, soundReference, R.embed);

            var vmlDrawingParts = oldContentPart switch
            {
                ChartsheetPart part => part.VmlDrawingParts,
                DialogsheetPart part => part.VmlDrawingParts,
                HandoutMasterPart part => part.VmlDrawingParts,
                InternationalMacroSheetPart part => part.VmlDrawingParts,
                MacroSheetPart part => part.VmlDrawingParts,
                NotesMasterPart part => part.VmlDrawingParts,
                NotesSlidePart part => part.VmlDrawingParts,
                SlideLayoutPart part => part.VmlDrawingParts,
                SlideMasterPart part => part.VmlDrawingParts,
                SlidePart part => part.VmlDrawingParts,
                WorksheetPart part => part.VmlDrawingParts,
                _ => null,
            };

            if (vmlDrawingParts is not null)
            {
                // Transitional: Copy VML Drawing parts, implicit relationship
                foreach (var vmlPart in vmlDrawingParts)
                {
                    var newVmlPart = newContentPart switch
                    {
                        ChartsheetPart part => part.AddNewPart<VmlDrawingPart>(),
                        DialogsheetPart part => part.AddNewPart<VmlDrawingPart>(),
                        HandoutMasterPart part => part.AddNewPart<VmlDrawingPart>(),
                        InternationalMacroSheetPart part => part.AddNewPart<VmlDrawingPart>(),
                        MacroSheetPart part => part.AddNewPart<VmlDrawingPart>(),
                        NotesMasterPart part => part.AddNewPart<VmlDrawingPart>(),
                        NotesSlidePart part => part.AddNewPart<VmlDrawingPart>(),
                        SlideLayoutPart part => part.AddNewPart<VmlDrawingPart>(),
                        SlideMasterPart part => part.AddNewPart<VmlDrawingPart>(),
                        SlidePart part => part.AddNewPart<VmlDrawingPart>(),
                        WorksheetPart part => part.AddNewPart<VmlDrawingPart>(),
                        _ => null,
                    };

                    try
                    {
                        var xd = new XDocument(vmlPart.GetXDocument());
                        foreach (var item in xd.Descendants(O.ink))
                        {
                            if (item.Attribute("i") is { } attr)
                            {
                                var i = attr.Value;
                                i = i.Replace(" ", "\r\n");
                                attr.Value = i;
                            }
                        }
                        newVmlPart.PutXDocument(xd);

                        PBT.AddRelationships(vmlPart, newVmlPart, [newVmlPart.GetXDocument().Root]);
                        CopyRelatedPartsForContentParts(vmlPart, newVmlPart, [newVmlPart.GetXDocument().Root]);
                    }
                    catch (XmlException)
                    {
                        using var srcStream = vmlPart.GetStream();
                        using var dstStream = newVmlPart.GetStream(FileMode.Create, FileAccess.Write);
                        srcStream.CopyTo(dstStream);
                    }
                }
            }
        }
    }

    private void CopyRelatedImage(
        OpenXmlPart oldContentPart,
        OpenXmlPart newContentPart,
        XElement imageReference,
        XName attributeName
    )
    {
        // First look to see if this relId has already been added to the new document.
        // This is necessary for those parts that get processed with both old and new ids, such as the comments
        // part.  This is not necessary for parts such as the main document part, but this code won't malfunction
        // in that case.

        var relId = (string)imageReference.Attribute(attributeName);
        if (newContentPart.HasRelationship(relId))
            return;

        var oldPartIdPair = oldContentPart.Parts.FirstOrDefault(p => p.RelationshipId == relId);
        if (oldPartIdPair != default)
        {
            var oldPart = oldPartIdPair.OpenXmlPart as ImagePart;
            var temp = GetOrAddImageCopy(oldPart);
            if (temp.ImagePart is null)
            {
                var contentType = oldPart?.ContentType;
                var targetExtension = contentType switch
                {
                    "image/bmp" => ".bmp",
                    "image/gif" => ".gif",
                    "image/png" => ".png",
                    "image/tiff" => ".tiff",
                    "image/x-icon" => ".ico",
                    "image/x-pcx" => ".pcx",
                    "image/jpeg" => ".jpg",
                    "image/x-emf" => ".emf",
                    "image/x-wmf" => ".wmf",
                    "image/svg+xml" => ".svg",
                    _ => ".image",
                };
                newContentPart.Features.GetRequired<IPartExtensionFeature>().Register(contentType, targetExtension);

                var newPart = newContentPart switch
                {
                    ChartDrawingPart part => part.AddImagePart(contentType),
                    ChartPart part => part.AddImagePart(contentType),
                    ChartsheetPart part => part.AddImagePart(contentType),
                    DiagramDataPart part => part.AddImagePart(contentType),
                    DiagramLayoutDefinitionPart part => part.AddImagePart(contentType),
                    DiagramPersistLayoutPart part => part.AddImagePart(contentType),
                    DrawingsPart part => part.AddImagePart(contentType),
                    HandoutMasterPart part => part.AddImagePart(contentType),
                    NotesMasterPart part => part.AddImagePart(contentType),
                    NotesSlidePart part => part.AddImagePart(contentType),
                    RibbonAndBackstageCustomizationsPart part => part.AddImagePart(contentType),
                    RibbonExtensibilityPart part => part.AddImagePart(contentType),
                    SlideLayoutPart part => part.AddImagePart(contentType),
                    SlideMasterPart part => part.AddImagePart(contentType),
                    SlidePart part => part.AddImagePart(contentType),
                    ThemeOverridePart part => part.AddImagePart(contentType),
                    ThemePart part => part.AddImagePart(contentType),
                    VmlDrawingPart part => part.AddImagePart(contentType),
                    WorksheetPart part => part.AddImagePart(contentType),
                    _ => null,
                };

                temp.ImagePart = newPart;
                var id = newContentPart.GetIdOfPart(newPart);
                temp.AddContentPartRelTypeResourceIdTupple(newContentPart, newPart.RelationshipType, id);

                using (var stream = oldPart.GetStream())
                    newPart.FeedData(stream);
                imageReference.SetAttributeValue(attributeName, id);
            }
            else
            {
                var refRel = newContentPart.DataPartReferenceRelationships.FirstOrDefault(rr =>
                    temp.ContentPartRelTypeIdList.Exists(cpr =>
                        cpr.ContentPart == newContentPart && cpr.RelationshipId == rr.Id
                    )
                );
                if (refRel is not null)
                {
                    var relationshipId = temp
                        .ContentPartRelTypeIdList.First(cpr =>
                            cpr.ContentPart == newContentPart && cpr.RelationshipId == refRel.Id
                        )
                        .RelationshipId;
                    imageReference.SetAttributeValue(attributeName, relationshipId);
                    return;
                }

                var cpr2 = temp.ContentPartRelTypeIdList.FirstOrDefault(c => c.ContentPart == newContentPart);
                if (cpr2 is not null)
                {
                    imageReference.SetAttributeValue(attributeName, cpr2.RelationshipId);
                }
                else
                {
                    var imagePart = (ImagePart)temp.ImagePart;
                    var existingImagePart = newContentPart.AddPart(imagePart);
                    var newId = newContentPart.GetIdOfPart(existingImagePart);
                    temp.AddContentPartRelTypeResourceIdTupple(newContentPart, imagePart.RelationshipType, newId);
                    imageReference.SetAttributeValue(attributeName, newId);
                }
            }
        }
        else
        {
            var er = oldContentPart.ExternalRelationships.FirstOrDefault(r => r.Id == relId);
            if (er is not null)
            {
                var newEr = newContentPart.AddExternalRelationship(er.RelationshipType, er.Uri);
                imageReference.SetAttributeValue(R.id, newEr.Id);
            }
            else
            {
                var newPart = newContentPart
                    .OpenXmlPackage.GetPackage()
                    .GetParts()
                    .FirstOrDefault(p => p.Uri == newContentPart.Uri);
                if (newPart is not null && !newPart.Relationships.Contains(relId))
                {
                    newPart.Relationships.Create(
                        new Uri("NULL", UriKind.RelativeOrAbsolute),
                        System.IO.Packaging.TargetMode.Internal,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        relId
                    );
                }
            }
        }
    }

    private void CopyRelatedMedia(
        OpenXmlPart oldContentPart,
        OpenXmlPart newContentPart,
        XElement imageReference,
        XName attributeName,
        string mediaRelationshipType
    )
    {
        var relId = (string)imageReference.Attribute(attributeName);
        if (string.IsNullOrEmpty(relId) || newContentPart.DataPartReferenceRelationships.Any(dpr => dpr.Id == relId)) // First look to see if this relId has already been added to the new document.
            return;

        var oldRel = oldContentPart.DataPartReferenceRelationships.FirstOrDefault(dpr => dpr.Id == relId);
        if (oldRel is null)
            return;

        var oldPart = oldRel.DataPart;
        var temp = GetOrAddMediaCopy(oldPart);
        if (temp.DataPart is null)
        {
            var ct = oldPart.ContentType;
            var ext = Path.GetExtension(oldPart.Uri.OriginalString);
            var newPart = newContentPart.OpenXmlPackage.CreateMediaDataPart(ct, ext);
            using (var stream = oldPart.GetStream())
                newPart.FeedData(stream);
            string id = null;
            string relationshipType = null;

            switch (mediaRelationshipType)
            {
                case "media":
                {
                    var mrr = newContentPart switch
                    {
                        SlidePart part => part.AddMediaReferenceRelationship(newPart),
                        SlideLayoutPart part => part.AddMediaReferenceRelationship(newPart),
                        SlideMasterPart part => part.AddMediaReferenceRelationship(newPart),
                        _ => null,
                    };

                    id = mrr?.Id;
                    relationshipType = "http://schemas.microsoft.com/office/2007/relationships/media";
                    break;
                }
                case "video":
                {
                    var vrr = newContentPart switch
                    {
                        SlidePart part => part.AddVideoReferenceRelationship(newPart),
                        HandoutMasterPart part => part.AddVideoReferenceRelationship(newPart),
                        NotesMasterPart part => part.AddVideoReferenceRelationship(newPart),
                        NotesSlidePart part => part.AddVideoReferenceRelationship(newPart),
                        SlideLayoutPart part => part.AddVideoReferenceRelationship(newPart),
                        SlideMasterPart part => part.AddVideoReferenceRelationship(newPart),
                        _ => null,
                    };

                    id = vrr?.Id;
                    relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video";
                    break;
                }
            }
            temp.DataPart = newPart;
            temp.AddContentPartRelTypeResourceIdTupple(newContentPart, relationshipType, id);
            imageReference.Attribute(attributeName).Set(id);
        }
        else
        {
            var desiredRelType = mediaRelationshipType switch
            {
                "media" => "http://schemas.microsoft.com/office/2007/relationships/media",
                "video" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video",
                _ => null,
            };

            var existingRel = temp.ContentPartRelTypeIdList.FirstOrDefault(cp =>
                cp.ContentPart == newContentPart && cp.RelationshipType == desiredRelType
            );
            if (existingRel is not null)
            {
                imageReference.Attribute(attributeName).Set(existingRel.RelationshipId);
            }
            else
            {
                var newPart = (MediaDataPart)temp.DataPart;
                string id = null;
                string relationshipType = null;
                switch (mediaRelationshipType)
                {
                    case "media":
                    {
                        var mrr = newContentPart switch
                        {
                            SlidePart part => part.AddMediaReferenceRelationship(newPart),
                            SlideLayoutPart part => part.AddMediaReferenceRelationship(newPart),
                            SlideMasterPart part => part.AddMediaReferenceRelationship(newPart),
                            _ => null,
                        };

                        id = mrr?.Id;
                        relationshipType = mrr?.RelationshipType;
                        break;
                    }
                    case "video":
                    {
                        var vrr = newContentPart switch
                        {
                            SlidePart part => part.AddVideoReferenceRelationship(newPart),
                            HandoutMasterPart part => part.AddVideoReferenceRelationship(newPart),
                            NotesMasterPart part => part.AddVideoReferenceRelationship(newPart),
                            NotesSlidePart part => part.AddVideoReferenceRelationship(newPart),
                            SlideLayoutPart part => part.AddVideoReferenceRelationship(newPart),
                            SlideMasterPart part => part.AddVideoReferenceRelationship(newPart),
                            _ => null,
                        };

                        id = vrr?.Id;
                        relationshipType = vrr?.RelationshipType;
                        break;
                    }
                }
                temp.AddContentPartRelTypeResourceIdTupple(newContentPart, relationshipType, id);
                imageReference.Attribute(attributeName).Set(id);
            }
        }
    }

    private ThemePart CopyThemePart(SlideMasterPart slideMasterPart, ThemePart oldThemePart, double scaleFactor)
    {
        var newThemePart = slideMasterPart.AddNewPart<ThemePart>();
        var newThemeDoc = new XDocument(oldThemePart.GetXDocument());
        SlideLayoutData.ScaleShapes(newThemeDoc, scaleFactor);
        newThemePart.PutXDocument(newThemeDoc);

        PBT.AddRelationships(oldThemePart, newThemePart, [newThemePart.GetXDocument().Root]);
        CopyRelatedPartsForContentParts(oldThemePart, newThemePart, [newThemePart.GetXDocument().Root]);

        if (_newDocument.PresentationPart.ThemePart is null)
            newThemePart = _newDocument.PresentationPart.AddPart(newThemePart);

        return newThemePart;
    }

    private SlideMasterPart CopySlideMasterPart(SlideMasterPart oldMasterPart, double scaleFactor)
    {
        var newMaster = _newDocument.PresentationPart.AddNewPart<SlideMasterPart>();

        // Add to presentation slide master list, need newID for layout IDs also
        var presentationPartDoc = _newDocument.PresentationPart.GetXDocument();
        presentationPartDoc
            .Root.Element(P.sldMasterIdLst)
            .Add(
                new XElement(
                    P.sldMasterId,
                    new XAttribute(NoNamespace.id, GetNextFreeId().ToString()),
                    new XAttribute(R.id, _newDocument.PresentationPart.GetIdOfPart(newMaster))
                )
            );

        // Ensure that master does not keep ids of old layouts
        var newMasterDoc = new XDocument(oldMasterPart.GetXDocument());
        var sldLayoutIdLst = newMasterDoc.Root.Element(P.sldLayoutIdLst);
        if (sldLayoutIdLst is null)
        {
            InsertElementInSchemaOrder(newMasterDoc.Root, new XElement(P.sldLayoutIdLst), PBT.OrderSlideMaster);
        }
        else
        {
            sldLayoutIdLst.Descendants(P.sldLayoutId).ToList().ForEach(e => e.Remove());
        }

        SlideLayoutData.ScaleShapes(newMasterDoc, scaleFactor);
        newMaster.PutXDocument(newMasterDoc);

        PBT.AddRelationships(oldMasterPart, newMaster, [newMaster.GetXDocument().Root]);
        CopyRelatedPartsForContentParts(oldMasterPart, newMaster, [newMaster.GetXDocument().Root]);

        _ = CopyThemePart(newMaster, oldMasterPart.ThemePart, scaleFactor);

        return newMaster;
    }

    private SlideLayoutPart CopySlideLayoutPart(
        SlideMasterPart newSlideMasterPart,
        SlideLayoutPart oldSlideLayoutPart,
        double scaleFactor
    )
    {
        var newLayout = newSlideMasterPart.AddNewPart<SlideLayoutPart>();
        newLayout.AddPart(newSlideMasterPart);

        var newLayoutDoc = new XDocument(oldSlideLayoutPart.GetXDocument());
        SlideLayoutData.ScaleShapes(newLayoutDoc, scaleFactor);
        newLayout.PutXDocument(newLayoutDoc);

        PBT.AddRelationships(oldSlideLayoutPart, newLayout, [newLayout.GetXDocument().Root]);
        CopyRelatedPartsForContentParts(oldSlideLayoutPart, newLayout, [newLayout.GetXDocument().Root]);

        var newMasterDoc = newSlideMasterPart.GetXDocument();
        newMasterDoc
            .Root.Element(P.sldLayoutIdLst)
            .Add(
                new XElement(
                    P.sldLayoutId,
                    new XAttribute(NoNamespace.id, GetNextFreeId().ToString()),
                    new XAttribute(R.id, newSlideMasterPart.GetIdOfPart(newLayout))
                )
            );

        return newLayout;
    }

    /// <summary>
    /// Inserts an element into the parent in the correct position according to the schema order.
    /// </summary>
    /// <param name="parent">The parent element to insert into</param>
    /// <param name="elementToInsert">The element to insert</param>
    /// <param name="order">Dictionary defining the correct element order</param>
    /// <remarks> See <see href="https://www.datypic.com/sc/ooxml/e-p_sldMaster.html">Open XML schema</see> for the order of elements.</remarks>
    private static void InsertElementInSchemaOrder(
        XElement parent,
        XElement elementToInsert,
        Dictionary<XName, int> order
    )
    {
        var targetOrder = order.GetValueOrDefault(elementToInsert.Name, 999);

        // Find the last element that should come before our element
        var insertAfter = parent.Elements().LastOrDefault(e => order.GetValueOrDefault(e.Name, 999) < targetOrder);

        if (insertAfter != null)
        {
            // Insert after the found element
            insertAfter.AddAfterSelf(elementToInsert);
        }
        else
        {
            // No elements come before this one, insert as first child
            parent.AddFirst(elementToInsert);
        }
    }
}
