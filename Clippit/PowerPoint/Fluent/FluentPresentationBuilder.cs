using System;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PBT = Clippit.PowerPoint.Fluent.PresentationBuilderTools;

namespace Clippit.PowerPoint.Fluent;

internal sealed partial class FluentPresentationBuilder : IFluentPresentationBuilder
{
    private readonly PresentationDocument _newDocument;
    private bool _isDocumentInitialized;

    internal FluentPresentationBuilder(PresentationDocument presentationDocument)
    {
        _newDocument = presentationDocument ?? throw new NullReferenceException(nameof(presentationDocument));

        if (_newDocument.PresentationPart is null)
        {
            _newDocument.AddPresentationPart();
            var xDocument = GetPresentationXDocument();
            _newDocument.PresentationPart.PutXDocument(xDocument);
        }

        _isDocumentInitialized = false;
        InitializeCaches();
    }

    private static XDocument GetPresentationXDocument()
    {
        XNamespace ns = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace drawingns = "http://schemas.openxmlformats.org/drawingml/2006/main";

        return new XDocument(
            new XElement(
                ns + "presentation",
                new XAttribute(XNamespace.Xmlns + "a", drawingns),
                new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                new XAttribute(XNamespace.Xmlns + "p", ns),
                new XElement(ns + "sldMasterIdLst"),
                new XElement(ns + "sldIdLst"),
                new XElement(ns + "notesSz", new XAttribute("cx", "6858000"), new XAttribute("cy", "9144000"))
            )
        )
        {
            Declaration = new XDeclaration("1.0", "UTF-8", "yes"),
        };
    }

    public void Dispose() => SaveAndCleanup();

    private void SaveAndCleanup()
    {
        // Remove sections list (all slides added to default section)
        var presentationDocument = _newDocument.PresentationPart.GetXDocument();
        var sectionLists = presentationDocument.Descendants(P14.sectionLst).ToList();
        foreach (var sectionList in sectionLists)
        {
            sectionList.Parent?.Remove(); // <p:ext> element
        }

        // Remove custom properties (source doc metadata irrelevant for generated document)
        var customPropsDocument = _newDocument.CustomFilePropertiesPart?.GetXDocument();
        if (customPropsDocument?.Root?.HasElements == true)
        {
            customPropsDocument.Root?.RemoveNodes();
        }

        foreach (var part in _newDocument.GetAllParts())
        {
            if (part.ContentType.EndsWith("+xml"))
            {
                var xd = part.GetXDocument();
                xd.Descendants().Attributes("smtClean").Remove();
                part.PutXDocument();
            }
            else if (part.Annotation<XDocument>() is not null)
                part.PutXDocument();
        }
    }

    public SlideMasterPart AddSlideMasterPart(SlideMasterPart slideMasterPart)
    {
        var sourceDocument = (PresentationDocument)slideMasterPart.OpenXmlPackage;
        EnsureDocumentInitialized(sourceDocument);

        var scaleFactor = GetScaleFactor(sourceDocument);
        var slideMasterData = GetOrAddSlideMasterPart(sourceDocument, slideMasterPart, scaleFactor);

        foreach (var slideLayoutPart in slideMasterPart.SlideLayoutParts)
        {
            _ = GetOrAddSlideLayoutPart(sourceDocument, slideLayoutPart, scaleFactor);
        }

        return slideMasterData.Part;
    }

    private void EnsureDocumentInitialized(PresentationDocument sourceDocument)
    {
        if (_isDocumentInitialized)
            return;

        CopyStartingParts(sourceDocument);
        CopyPresentationParts(sourceDocument);

        _slideSize = sourceDocument.PresentationPart.Presentation.SlideSize.CloneNode(true) as SlideSize;

        var newPresentation = _newDocument.PresentationPart.GetXDocument();
        if (newPresentation.Root.Element(P.sldIdLst) is null)
        {
            newPresentation.Root.Add(new XElement(P.sldIdLst));
        }

        _isDocumentInitialized = true;
    }

    public SlidePart AddSlidePart(SlidePart slidePart)
    {
        var sourceDocument = (PresentationDocument)slidePart.OpenXmlPackage;
        ArgumentNullException.ThrowIfNull(sourceDocument);

        try
        {
            return AddSlidePartImpl(sourceDocument, slidePart);
        }
        catch (PresentationBuilderInternalException dbie)
        {
            if (dbie.Message.Contains("{0}"))
                throw new PresentationBuilderException(string.Format(dbie.Message, slidePart.Uri));
            throw;
        }
    }

    private SlidePart AddSlidePartImpl(PresentationDocument sourceDocument, SlidePart slidePart)
    {
        EnsureDocumentInitialized(sourceDocument);

        var scaleFactor = GetScaleFactor(sourceDocument);

        // TODO: Maintain it globally on the builder level, instead of calculating it for each slide add operation
        var newPresentation = _newDocument.PresentationPart.GetXDocument();
        uint newId = 256;
        var ids = newPresentation.Root.Descendants(P.sldId).Select(f => (uint)f.Attribute(NoNamespace.id)).ToList();
        if (ids.Count != 0)
            newId = ids.Max() + 1;

        var newSlide = _newDocument.PresentationPart.AddNewPart<SlidePart>();
        using (var sourceStream = slidePart.GetStream())
        {
            newSlide.FeedData(sourceStream);
        }

        var slideDocument = newSlide.GetXDocument();
        SlideLayoutData.ScaleShapes(slideDocument, scaleFactor);

        PBT.AddRelationships(slidePart, newSlide, [newSlide.GetXDocument().Root]);
        CopyRelatedPartsForContentParts(slidePart, newSlide, [newSlide.GetXDocument().Root]);
        CopyTableStyles(sourceDocument, newSlide);

        if (slidePart.NotesSlidePart is { } notesSlide)
        {
            if (_newDocument.PresentationPart.NotesMasterPart is null)
                CopyNotesMaster(sourceDocument);
            var newPart = newSlide.AddNewPart<NotesSlidePart>();
            newPart.PutXDocument(notesSlide.GetXDocument());
            newPart.AddPart(newSlide);
            if (_newDocument.PresentationPart.NotesMasterPart is not null)
                newPart.AddPart(_newDocument.PresentationPart.NotesMasterPart);
            PBT.AddRelationships(notesSlide, newPart, [newPart.GetXDocument().Root]);
            CopyRelatedPartsForContentParts(slidePart.NotesSlidePart, newPart, [newPart.GetXDocument().Root]);
        }

        var slideLayoutData = GetOrAddSlideLayoutPart(sourceDocument, slidePart.SlideLayoutPart, scaleFactor);
        newSlide.AddPart(slideLayoutData.Part);

        if (slidePart.SlideCommentsPart is not null)
            CopyComments(sourceDocument, slidePart, newSlide);

        newPresentation = _newDocument.PresentationPart.GetXDocument();
        newPresentation
            .Root.Element(P.sldIdLst)
            .Add(
                new XElement(
                    P.sldId,
                    new XAttribute(NoNamespace.id, newId.ToString()),
                    new XAttribute(R.id, _newDocument.PresentationPart.GetIdOfPart(newSlide))
                )
            );

        return newSlide;
    }
}
