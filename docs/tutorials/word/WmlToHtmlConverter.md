---
uid: Tutorial.Word.WmlToHtmlConverter
---
# Convert Word to HTML

Namespace: `Clippit.Word`

Convert a Word document to an HTML `XElement`, with configurable CSS generation and image handling.

```csharp
public static class WmlToHtmlConverter {
    public static XElement ConvertToHtml(
        WmlDocument doc, WmlToHtmlConverterSettings htmlConverterSettings)
    {...}

    public static XElement ConvertToHtml(
        WordprocessingDocument wordDoc, WmlToHtmlConverterSettings htmlConverterSettings)
    {...}
}
```

The converter produces a complete HTML document as an `XElement` (XHTML). It generates CSS classes
for paragraph and character styles, handles numbering/lists, and processes images through a
configurable `ImageHandler` callback.

An extension method is also available directly on `WmlDocument`:

```csharp
WmlDocument doc = new WmlDocument("input.docx");
XElement html = doc.ConvertToHtml(settings);
```

### WmlToHtmlConverterSettings

| Field | Type | Default |
|---|---|---|
| `PageTitle` | `string` | `""` |
| `CssClassPrefix` | `string` | `"pt-"` |
| `FabricateCssClasses` | `bool` | `true` |
| `GeneralCss` | `string` | `"span { white-space: pre-wrap; }"` |
| `AdditionalCss` | `string` | `""` |
| `RestrictToSupportedLanguages` | `bool` | `false` |
| `RestrictToSupportedNumberingFormats` | `bool` | `false` |
| `ImageHandler` | `Func<ImageInfo, XElement>` | `null` |

### ImageInfo

The `ImageHandler` callback receives an `ImageInfo` object for each image in the document:

| Field | Type | Description |
|---|---|---|
| `Image` | `SixLabors.ImageSharp.Image` | The decoded image |
| `ImgStyleAttribute` | `XAttribute` | The computed `style` attribute (width/height) |
| `ContentType` | `string` | The image MIME type |
| `DrawingElement` | `XElement` | The source OpenXml drawing element |
| `AltText` | `string` | Alternative text for the image |

### WmlToHtmlConverter Sample

```csharp
var doc = new WmlDocument("input.docx");

var settings = new WmlToHtmlConverterSettings
{
    PageTitle = "My Document",
    CssClassPrefix = "doc-",
    FabricateCssClasses = true,
    AdditionalCss = "body { font-family: Calibri, sans-serif; }",
    ImageHandler = imageInfo =>
    {
        // Convert images to inline base64 data URIs
        using var stream = new MemoryStream();
        imageInfo.Image.SaveAsPng(stream);
        var base64 = Convert.ToBase64String(stream.ToArray());
        var imgElement = new XElement(
            Xhtml.img,
            imageInfo.ImgStyleAttribute,
            new XAttribute("src", $"data:image/png;base64,{base64}"),
            new XAttribute("alt", imageInfo.AltText ?? "")
        );
        return imgElement;
    }
};

var html = WmlToHtmlConverter.ConvertToHtml(doc, settings);
File.WriteAllText("output.html", html.ToString());
```
