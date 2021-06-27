# Build Presentation

Namespace: `Clippit.PowerPoint`

Combine collection of `SlideSource`'s into one PowerPoint presentations.

```csharp
public static class PresentationBuilder {
    public static PmlDocument BuildPresentation(List<SlideSource> sources)
    {...}
}
```

Original behavior is documented at Eric's blog:

http://www.ericwhite.com/blog/presentationbuilder-developer-center/

#### Major changes

- __Removed Custom Properties__. Clippit does not copy document properties from the first `SlideSource` into result document.
- __Removed Section List__. Clippit copy all slides into Default section and does not copy list of sections from first `SlideSource`.
- __Structural comparison for Theme, Master, Layout__. Clippit structurally compare Themes, Masters and Layout and maintain minimal amount of these parts in target presentation to guarantee no visual artifacts on slides in the target presentation.
- __Shape auto-scale__. Clippit automatically scale all shapes on the slides when you merge slides of different size.

#### Fixes

- Appropriate file extentions generated for ImageParts based on ContentType (instead of `*.bin` for all images)
- Fixed multiple write access to the same part
- Optimized Stream lifetime and GC pressure


### BuildPresentation Sample

```csharp {highlight:[6]}
var sources = new List<SlideSource>()
{
    new SlideSource(new PmlDocument(file1), start:0, count:1, keepMaster:true),
    new SlideSource(new PmlDocument(file2), start:9, count:3, keepMaster:false)
};
PresentationBuilder.BuildPresentation(sources)
    .SaveAs(resultFile);
```

This code combines slides from two PowerPoint presentations:
- It takes Title slide from `file1` and copy master using by first slide (with all layouts)
- Then copy three slides from `file2` starting from slide 10 without master (reusing master and layouts moved from `file1`) 