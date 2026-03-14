---
uid: Tutorial.PowerPoint.BuildPresentation.FluentApi
---
# Fluent Presentation Builder

Namespace: `Clippit.PowerPoint`

Build PowerPoint presentations incrementally by adding individual slide master parts and slide parts
using a fluent API pattern.

```csharp
public static partial class PresentationBuilder {
    public static IFluentPresentationBuilder Create(PresentationDocument document)
    {...}

    public static PresentationDocument NewDocument(Stream stream)
    {...}
}

public interface IFluentPresentationBuilder : IDisposable {
    SlideMasterPart AddSlideMasterPart(SlideMasterPart slideMasterPart);
    SlidePart AddSlidePart(SlidePart slidePart);
}
```

Unlike `BuildPresentation` which combines `SlideSource` collections in a single call,
the fluent API allows you to build a presentation by adding parts one at a time.
This is useful when slides come from different sources or are assembled dynamically.

#### Key Features

- **Incremental construction** -- add slides one at a time as they become available
- **Part-level control** -- add slide masters and slides independently
- **Structural deduplication** -- themes, masters, and layouts are structurally compared
  and deduplicated, maintaining minimal parts in the target presentation
- **Shape auto-scaling** -- shapes are automatically scaled when merging slides of different sizes

### FluentPresentationBuilder Sample

```csharp
using var stream = new MemoryStream();
using var targetDoc = PresentationBuilder.NewDocument(stream);
using var builder = PresentationBuilder.Create(targetDoc);

// Open source presentations
using var source1 = PresentationDocument.Open("presentation1.pptx", false);
using var source2 = PresentationDocument.Open("presentation2.pptx", false);

// Add a slide master from the first presentation
var masterPart = source1.PresentationPart.SlideMasterParts.First();
builder.AddSlideMasterPart(masterPart);

// Add individual slides from different sources
foreach (var slidePart in source1.PresentationPart.SlideParts.Take(3))
{
    builder.AddSlidePart(slidePart);
}

foreach (var slidePart in source2.PresentationPart.SlideParts)
{
    builder.AddSlidePart(slidePart);
}

// Save the result
targetDoc.Save();
var result = new PmlDocument("result.pptx", stream.ToArray());
result.SaveAs("combined.pptx");
```
