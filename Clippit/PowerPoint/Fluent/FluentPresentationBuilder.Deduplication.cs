using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;

namespace Clippit.PowerPoint.Fluent;

internal partial class FluentPresentationBuilder
{
    private readonly List<ContentData> _mediaCache = [];
    private readonly Dictionary<SlideMasterPart, SlideMasterData> _slideMasters = [];
    private SlideSize _slideSize;

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
        return GetOrAddCachedMedia(new ImageData(oldImage));
    }

    // General function for handling media that tries to use an existing media item if they are the same
    private MediaData GetOrAddMediaCopy(DataPart oldMedia)
    {
        return GetOrAddCachedMedia(new MediaData(oldMedia));
    }

    private T GetOrAddCachedMedia<T>(T contentData)
        where T : ContentData
    {
        var duplicateItem = _mediaCache.FirstOrDefault(x => x.Compare(contentData));
        if (duplicateItem != null)
        {
            return (T)duplicateItem;
        }

        _mediaCache.Add(contentData);
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
