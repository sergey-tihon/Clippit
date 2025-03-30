using System;
using System.Linq;

namespace Clippit.PowerPoint.Fluent;

internal sealed partial class FluentPresentationBuilder
{
    private uint GetNextFreeId()
    {
        uint newId = 0;

        var presentationPartDoc = _newDocument.PresentationPart.GetXDocument();
        var masterIds = presentationPartDoc
            .Root.Descendants(P.sldMasterId)
            .Select(f => (uint)f.Attribute(NoNamespace.id))
            .ToList();
        if (masterIds.Count != 0)
            newId = Math.Max(newId, masterIds.Max());

        foreach (var slideMasterData in _slideMasters.Values)
        {
            var masterPartDoc = slideMasterData.Part.GetXDocument();
            var layoutIds = masterPartDoc
                .Root.Descendants(P.sldLayoutId)
                .Select(f => (uint)f.Attribute(NoNamespace.id))
                .ToList();
            if (layoutIds.Count != 0)
                newId = Math.Max(newId, layoutIds.Max());
        }

        return newId == 0 ? 2147483648 : newId + 1;
    }
}
