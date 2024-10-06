// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Clippit.Excel;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.PowerPoint
{
    public class SlideSource
    {
        public PmlDocument PmlDocument { get; set; }
        public int Start { get; set; }
        public int Count { get; set; }
        public bool KeepMaster { get; set; }

        public SlideSource(PmlDocument source, bool keepMaster)
        {
            PmlDocument = source;
            Start = 0;
            Count = int.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(string fileName, bool keepMaster)
        {
            PmlDocument = new PmlDocument(fileName);
            Start = 0;
            Count = int.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(PmlDocument source, int start, bool keepMaster)
        {
            PmlDocument = source;
            Start = start;
            Count = int.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(string fileName, int start, bool keepMaster)
        {
            PmlDocument = new PmlDocument(fileName);
            Start = start;
            Count = int.MaxValue;
            KeepMaster = keepMaster;
        }

        public SlideSource(PmlDocument source, int start, int count, bool keepMaster)
        {
            PmlDocument = source;
            Start = start;
            Count = count;
            KeepMaster = keepMaster;
        }

        public SlideSource(string fileName, int start, int count, bool keepMaster)
        {
            PmlDocument = new PmlDocument(fileName);
            Start = start;
            Count = count;
            KeepMaster = keepMaster;
        }
    }

    public static partial class PresentationBuilder
    {
        public static IFluentPresentationBuilder Create(PresentationDocument document)
        {
            return new FluentPresentationBuilder(document);
        }

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

        public static IList<PmlDocument> PublishSlides(PmlDocument src)
        {
            using var streamSrcDoc = new OpenXmlMemoryStreamDocument(src);
            using var srcDoc = streamSrcDoc.GetPresentationDocument(new OpenSettings { AutoSave = false });
            return PublishSlides(srcDoc, src.FileName).ToList();
        }

        public static IEnumerable<PmlDocument> PublishSlides(PresentationDocument srcDoc, string fileName)
        {
            var slidesIds = GetSlidesInOrder(srcDoc);
            var slideNameRegex = SlideNameRegex();
            for (var slideNumber = 0; slideNumber < slidesIds.Count; slideNumber++)
            {
                var srcSlidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slidesIds[slideNumber]);

                using var streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
                using (var output = streamDoc.GetPresentationDocument(new OpenSettings { AutoSave = false }))
                {
                    ExtractSlide(srcSlidePart, output);
                }

                var slideDoc = streamDoc.GetModifiedPmlDocument();
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    slideDoc.FileName = slideNameRegex.Replace(fileName, $"_{slideNumber + 1:000}.pptx");
                }

                yield return slideDoc;
            }
        }

        private static List<string> GetSlidesInOrder(PresentationDocument srcDoc)
        {
            return srcDoc
                .PresentationPart.GetXElement()
                .Descendants(P.sldId)
                .Select(x => x.Attribute(R.id)!.Value)
                .ToList();
        }

        private static void ExtractSlide(SlidePart slidePart, PresentationDocument output)
        {
            using var builder = new FluentPresentationBuilder(output);
            try
            {
                var newSlidePart = builder.AddSlide(slidePart);

                // Remove the show attribute from the slide element (if it exists)
                var slideDocument = newSlidePart.GetXDocument();
                slideDocument.Root?.Attribute(NoNamespace.show)?.Remove();

                // Set the title of the new presentation to the title of the slide
                var title = PresentationBuilderTools.GetSlideTitle(newSlidePart.GetXElement());
                output.PackageProperties.Title = title;
            }
            catch (PresentationBuilderInternalException dbie)
            {
                if (dbie.Message.Contains("{0}"))
                    throw new PresentationBuilderException(string.Format(dbie.Message, slidePart.Uri));
                throw;
            }
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
                            builder.AddSlideMaster(slideMasterPart);
                        }
                    }

                    var slideIds = GetSlidesInOrder(doc);
                    var (count, start) = (source.Count, source.Start);
                    while (count > 0 && start < slideIds.Count)
                    {
                        var slidePart = (SlidePart)doc.PresentationPart.GetPartById(slideIds[start]);
                        builder.AddSlide(slidePart);

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

        [GeneratedRegex(".pptx", RegexOptions.IgnoreCase, "en-US")]
        private static partial Regex SlideNameRegex();
    }
}
