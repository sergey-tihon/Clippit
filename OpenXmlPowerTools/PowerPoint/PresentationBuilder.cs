// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
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

    public static class PresentationBuilder
    {
        public static PmlDocument BuildPresentation(List<SlideSource> sources)
        {
            using var streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
            using (var output = streamDoc.GetPresentationDocument(new OpenSettings { AutoSave = false}))
            {
                BuildPresentation(sources, output);
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
            var slideList = srcDoc.PresentationPart.GetXDocument().Root.Descendants(P.sldId).ToList();
            for (var slideNumber = 0; slideNumber < slideList.Count; slideNumber++)
            {
                using var streamDoc = OpenXmlMemoryStreamDocument.CreatePresentationDocument();
                using (var output = streamDoc.GetPresentationDocument(new OpenSettings { AutoSave = false}))
                {
                    ExtractSlide(srcDoc, slideNumber, output);

                    var slidePartId = slideList.ElementAt(slideNumber).Attribute(R.id)?.Value;
                    var slidePart = (SlidePart)srcDoc.PresentationPart.GetPartById(slidePartId);
                    output.PackageProperties.Title = PresentationBuilderTools.GetSlideTitle(slidePart);
                }

                var slideDoc = streamDoc.GetModifiedPmlDocument();
                if (!string.IsNullOrWhiteSpace(fileName))
                {
                    slideDoc.FileName =
                        Regex.Replace(fileName, ".pptx", $"_{slideNumber + 1:000}.pptx", RegexOptions.IgnoreCase);
                }

                yield return slideDoc;
            }
        }

        private static void ExtractSlide(PresentationDocument srcDoc, int slideNumber, PresentationDocument output)
        {
            using var fluentBuilder = new FluentPresentationBuilder(output);
            try
            {
                fluentBuilder.AppendSlides(srcDoc, slideNumber, 1, true);
            }
            catch (PresentationBuilderInternalException dbie)
            {
                if (dbie.Message.Contains("{0}"))
                    throw new PresentationBuilderException(string.Format(dbie.Message, slideNumber));
                throw;
            }
        }

        private static void BuildPresentation(List<SlideSource> sources, PresentationDocument output)
        {
            using var fluentBuilder = new FluentPresentationBuilder(output);
            
            var sourceNum = 0;
            var openSettings = new OpenSettings {AutoSave = false};
            foreach (var source in sources)
            {
                using var streamDoc = new OpenXmlMemoryStreamDocument(source.PmlDocument);
                using var doc = streamDoc.GetPresentationDocument(openSettings);
                try
                {
                    fluentBuilder.AppendSlides(doc, source.Start, source.Count, source.KeepMaster);
                }
                catch (PresentationBuilderInternalException dbie)
                {
                    if (dbie.Message.Contains("{0}"))
                        throw new PresentationBuilderException(string.Format(dbie.Message, sourceNum));
                    throw dbie;
                }

                sourceNum++;
            }
        }
    }
}
