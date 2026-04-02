using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace Clippit.PowerPoint.Fluent;

internal partial class FluentPresentationBuilder
{
    private readonly Dictionary<ContentDataKey, ContentData> _mediaCache = [];

    // Identity-keyed caches: if the same OpenXmlPart instance is encountered again
    // (e.g. a 1 GB video embedded on multiple slides) we reuse the already-computed
    // ContentData directly, avoiding a second expensive stream read + SHA-256.
    private readonly Dictionary<ImagePart, ContentData> _imagePartCache = new(ReferenceEqualityComparer.Instance);
    private readonly Dictionary<DataPart, ContentData> _dataPartCache = new(ReferenceEqualityComparer.Instance);

    private readonly Dictionary<SlideMasterPart, SlideMasterData> _slideMasters = [];
    private SlideSize _slideSize;
    private uint _nextSlideId;

    private void InitializeCaches()
    {
        if (_newDocument.PresentationPart is not { } presentation)
            return;

        foreach (var slideMasterPart in presentation.SlideMasterParts)
        {
            foreach (var slideLayoutPart in slideMasterPart.SlideLayoutParts)
            {
                _ = GetOrAddSlideLayoutPart(_newDocument, slideLayoutPart, 1.0f);
            }
        }

        // TODO: enumerate all images, media, master and layouts
        _slideSize = presentation.Presentation.SlideSize;

        var existingSlideIds =
            presentation
                .GetXDocument()
                ?.Root?.Descendants(P.sldId)
                .Select(f => (uint)f.Attribute(NoNamespace.id))
                .ToList()
            ?? [];
        _nextSlideId = existingSlideIds.Count > 0 ? existingSlideIds.Max() + 1 : 256;
    }

    private double GetScaleFactor(PresentationDocument sourceDocument)
    {
        var slideSize = sourceDocument.PresentationPart.Presentation.SlideSize;
        var scaleFactorX = (double)_slideSize.Cx / slideSize.Cx;
        var scaleFactorY = (double)_slideSize.Cy / slideSize.Cy;
        return Math.Min(scaleFactorX, scaleFactorY);
    }

    // General function for handling images that tries to use an existing image if they are the same
    private ImageData GetOrAddImageCopy(ImagePart oldImage)
    {
        if (_imagePartCache.TryGetValue(oldImage, out var cached))
            return (ImageData)cached;

        var imageData = GetOrAddCachedMedia(new ImageData(oldImage));
        _imagePartCache[oldImage] = imageData;
        return imageData;
    }

    // General function for handling media that tries to use an existing media item if they are the same
    private MediaData GetOrAddMediaCopy(DataPart oldMedia)
    {
        if (_dataPartCache.TryGetValue(oldMedia, out var cached))
            return (MediaData)cached;

        var mediaData = GetOrAddCachedMedia(new MediaData(oldMedia));
        _dataPartCache[oldMedia] = mediaData;
        return mediaData;
    }

    private T GetOrAddCachedMedia<T>(T contentData)
        where T : ContentData
    {
        var key = contentData.Key;
        if (_mediaCache.TryGetValue(key, out var existing))
            return (T)existing;

        _mediaCache[key] = contentData;
        return contentData;
    }

    // General function for handling SlideMasterPart that tries to use an existing SlideMasterPart if they are the same
    private SlideMasterData GetOrAddSlideMasterPart(
        PresentationDocument presentationDocument,
        SlideMasterPart slideMasterPart,
        double scaleFactor
    )
    {
        if (_slideMasters.TryGetValue(slideMasterPart, out var slideMasterData))
        {
            return slideMasterData;
        }

        slideMasterData = new SlideMasterData(slideMasterPart, scaleFactor);
        foreach (var item in _slideMasters.Values)
        {
            if (item.CompareTo(slideMasterData) == 0)
                return item;
        }

        if (!ReferenceEquals(presentationDocument, _newDocument))
        {
            var newSlideMasterPart = CopySlideMasterPart(slideMasterPart, scaleFactor);
            slideMasterData = new SlideMasterData(newSlideMasterPart, 1.0);
        }

        _slideMasters.Add(slideMasterPart, slideMasterData);
        return slideMasterData;
    }

    // General function for handling SlideMasterPart that tries to use an existing SlideMasterPart if they are the same
    private SlideLayoutData GetOrAddSlideLayoutPart(
        PresentationDocument presentationDocument,
        SlideLayoutPart slideLayoutPart,
        double scaleFactor
    )
    {
        var slideMasterData = GetOrAddSlideMasterPart(
            presentationDocument,
            slideLayoutPart.SlideMasterPart,
            scaleFactor
        );

        if (slideMasterData.SlideLayouts.TryGetValue(slideLayoutPart, out var slideLayoutData))
        {
            return slideLayoutData;
        }

        slideLayoutData = new SlideLayoutData(slideLayoutPart, scaleFactor);
        foreach (var item in slideMasterData.SlideLayouts.Values)
        {
            if (item.CompareTo(slideLayoutData) == 0)
                return item;
        }

        if (!ReferenceEquals(presentationDocument, _newDocument))
        {
            var newSlideLayoutPart = CopySlideLayoutPart(slideMasterData.Part, slideLayoutPart, scaleFactor);
            slideLayoutData = new SlideLayoutData(newSlideLayoutPart, 1.0);
        }

        slideMasterData.SlideLayouts.Add(slideLayoutPart, slideLayoutData);
        return slideLayoutData;
    }
}
