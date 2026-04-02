---
uid: Tutorial.Word.DocumentAssembler.FitWithin
---
# FitWithin Image Sizing Mode

Namespace: `Clippit.Word`

## Introduction

The `Image` directive in `DocumentAssembler` supports a `FitWithin` attribute that controls how images are sized relative to their placeholder in the document template.

By default, `DocumentAssembler` always scales an image to fill the placeholder dimensions (maintaining aspect ratio). The `FitWithin="true"` attribute changes this behaviour to a more intuitive "fit, don't upscale" mode:

- **Small images** (natural size fits within the placeholder): rendered at their original pixel dimensions — no upscaling occurs.
- **Large images** (exceed the placeholder in either dimension): scaled down proportionally so they fit within the bounds.

This is useful when document templates receive images of varying sizes and upscaling small images would look blurry or pixelated.

## Usage

### Content Control Style

In a Word document template, use a content control (structured document tag) with the `FitWithin` attribute on the `Image` directive:

```xml
<Image Select="./Photo" FitWithin="true" />
```

## Accepted Values

The `FitWithin` attribute accepts XSD `xs:boolean` values (case-insensitive):

| Value   | Effect                                    |
| ------- | ----------------------------------------- |
| `true`  | Fit-within mode enabled (no upscaling)    |
| `1`     | Fit-within mode enabled (no upscaling)    |
| `false` | Default sizing (always scale to fill)     |
| `0`     | Default sizing (always scale to fill)     |

Omitting the attribute entirely is equivalent to `FitWithin="false"`.

## Example

Given a template with a 200×200 px placeholder:

### Small Image (50×50 px)

```xml
<Image Select="./Photo" FitWithin="true" />
```

**Result:** The image is kept at 50×50 px. It is not stretched to fill the 200×200 placeholder.

### Large Image (400×200 px)

**Result:** The image is scaled down proportionally. With a 200×200 placeholder, the scale factor is `min(200/400, 200/200) = 0.5`, giving a final size of 200×100 px.

## Comparison of Sizing Modes

| Mode | Attribute | Small image | Large image |
|---|---|---|---|
| Default | *(none)* | Upscaled to fill placeholder | Scaled to fill placeholder |
| Keep original size | `preferRelativeResize="0"` on template drawing | Kept at original size | May overflow placeholder |
| Fit within | `FitWithin="true"` | Kept at original size | Scaled down to fit |

## Priority

When `FitWithin="true"` is specified, it takes priority over the other sizing modes (`keepSourceImageAspect` and `keepOriginalImageSize`) that are derived from the template drawing's XML attributes.

## Further Reading

- For template testing and examples, see `DocumentAssemblerTests.cs` in the test project, particularly the `DA291_Image_FitWithin_*` and `DA292_Image_FitWithin_*` test methods.
- The `Image` directive reference describes all supported attributes.

Changes merged in: [#168](https://github.com/sergey-tihon/Clippit/pull/168)
