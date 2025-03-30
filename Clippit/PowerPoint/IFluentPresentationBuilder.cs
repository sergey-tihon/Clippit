using System;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint;

public interface IFluentPresentationBuilder : IDisposable
{
    public SlideMasterPart AddSlideMasterPart(SlideMasterPart slideMasterPart);
    public SlidePart AddSlidePart(SlidePart slidePart);
}
