# Slide Publishing

Namespace: `Clippit.PowerPoint`

Split PowerPoint presentation (`PmlDocument`) into lazy sequence of one-slide presentations.

```csharp
public static class PresentationBuilder {
    public static IEnumerable<PmlDocument> PublishSlides(PmlDocument src)
    {...}
}
```

This is fully managed alternative of [Presentation.PublishSlides](https://docs.microsoft.com/en-us/office/vba/api/powerpoint.presentation.publishslides) that does not require MS Office to be installed on the machine.

#### Key facts

1. All extracted slides are marked as visible (even if they were hidden in the source presentation)
1. Generated slides contain only one master with only one used layout (master will be renamed). Behavior is similar to `Presentation.PublishSlides` and allow to dramatically decrease total size of generated slides.
1. `PublishSlides` is up to `6x` times faster than `BuildPresentation` for the same task (Because we open source presentation only once)
1. Slide title promoted to generated presentation title (when layout has slide). Last modified date propagated from source document.

## Publishing sample

```csharp {highlight:[2]}
var presentation = new PmlDocument(sourcePath);
var slides = PresentationBuilder.PublishSlides(presentation)
foreach (var slide in slides)
{
    var targetPath = Path.Combine(targetDir, Path.GetFileName(slide.FileName))
    slide.SaveAs(targetPath);
}
```

## Composing slides to one presentation

You can combine generate slide back to one presentation without breaking them

```csharp {highlight:['4-5']}
var presentation = new PmlDocument(sourceFile);
var slides = PresentationBuilder.PublishSlides(presentation).ToList();

var sources = slides.Select(x => new SlideSource(x, keepMaster:true)).ToList();
PresentationBuilder.BuildPresentation(sources)
    .SaveAs(newFileName);
```

this code will generate presentation with multiple one-layout masters.

If you want to have full master inside your generated presentation the first slide source should carry this master

```csharp {highlight:['4-7', 14]}
var presentation = new PmlDocument(sourceFile);

// generate presentation with all masters
var onlyMasters = PresentationBuilder.BuildPresentation(
    new List<SlideSource> {
        new SlideSource(presentation, start:0, count:0, keepMaster:true)
    });

// publish slides with one-layout masters
var slides = PresentationBuilder.PublishSlides(presentation);

// compose them together using only master as the first SlideSource
var sources = new List<SlideSource> {
    new SlideSource(onlyMaster, keepMaster:true)};
sources.AddRange(slides.Select(x => new SlideSource(x, keepMaster:false)));
PresentationBuilder.BuildPresentation(sources)
    .SaveAs(newFileName);
```
