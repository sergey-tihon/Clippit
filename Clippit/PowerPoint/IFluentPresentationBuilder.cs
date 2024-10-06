using System;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint;

public interface IFluentPresentationBuilder : IDisposable
{
    public void AddSlideMaster(SlideMasterPart slideMasterPart);
    public SlidePart AddSlide(SlidePart slidePart);
}
