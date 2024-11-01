// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;
using PBT = Clippit.PowerPoint.Fluent.PresentationBuilderTools;

namespace Clippit.PowerPoint;

public static partial class PresentationBuilder
{
    public static PmlDocument BuildPresentation(List<SlideSource> sources)
    {
        using var streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
        using (var output = streamDoc.GetPresentationDocument(new OpenSettings { AutoSave = false }))
        {
            BuildPresentation(sources, output);
            output.PackageProperties.Modified = DateTime.Now;
        }
        return streamDoc.GetModifiedPmlDocument();
    }

    private static void BuildPresentation(List<SlideSource> sources, PresentationDocument output)
    {
        using var builder = Create(output);

        var sourceNum = 0;
        var openSettings = new OpenSettings { AutoSave = false };
        foreach (var source in sources)
        {
            using var streamDoc = new OpenXmlMemoryStreamDocument(source.PmlDocument);
            using var doc = streamDoc.GetPresentationDocument(openSettings);
            try
            {
                if (source.KeepMaster)
                {
                    foreach (var slideMasterPart in doc.PresentationPart.SlideMasterParts)
                    {
                        builder.AddSlideMasterPart(slideMasterPart);
                    }
                }

                var slideIds = PBT.GetSlideIdsInOrder(doc);
                var (count, start) = (source.Count, source.Start);
                while (count > 0 && start < slideIds.Count)
                {
                    var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slideIds[start]);
                    builder.AddSlidePart(slidePart);

                    start++;
                    count--;
                }
            }
            catch (PresentationBuilderInternalException dbie)
            {
                if (dbie.Message.Contains("{0}"))
                    throw new PresentationBuilderException(string.Format(dbie.Message, sourceNum));
                throw;
            }

            sourceNum++;
        }
    }
}
