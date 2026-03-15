---
uid: Tutorial.Word.HtmlToWmlConverter
---
# Convert HTML to Word

Namespace: `Clippit.Html`

Convert an XHTML document to a Word document, with configurable CSS cascading and page layout settings.

```csharp
public class HtmlToWmlConverter {
    public static WmlDocument ConvertHtmlToWml(
        string defaultCss, string authorCss, string userCss,
        XElement xhtml, HtmlToWmlConverterSettings settings)
    {...}

    public static WmlDocument ConvertHtmlToWml(
        string defaultCss, string authorCss, string userCss,
        XElement xhtml, HtmlToWmlConverterSettings settings,
        WmlDocument emptyDocument, string annotatedHtmlDumpFileName)
    {...}

    public static HtmlToWmlConverterSettings GetDefaultSettings()
    {...}

    public static string CleanUpCss(string css)
    {...}
}
```

The converter accepts XHTML as an `XElement` and produces a `WmlDocument`. CSS is applied in three
layers following the CSS cascade: default (browser-like defaults), author (document styles), and
user (overrides). The `GetDefaultSettings()` method returns a pre-configured
`HtmlToWmlConverterSettings` with sensible page layout defaults.

### HtmlToWmlConverterSettings

| Field | Type | Description |
|---|---|---|
| `MajorLatinFont` | `string` | Major (heading) Latin font |
| `MinorLatinFont` | `string` | Minor (body) Latin font |
| `DefaultFontSize` | `double` | Default font size |
| `DefaultSpacingElement` | `XElement` | Default paragraph spacing |
| `DefaultSpacingElementForParagraphsInTables` | `XElement` | Default spacing in tables |
| `SectPr` | `XElement` | Section properties (page size, margins) |
| `DefaultBlockContentMargin` | `string` | Default margin for block content |
| `BaseUriForImages` | `string` | Base URI for resolving image paths |

The settings also expose read-only properties derived from `SectPr`:
`PageWidthTwips`, `PageMarginLeftTwips`, `PageMarginRightTwips`,
`PageWidthEmus`, `PageMarginLeftEmus`, `PageMarginRightEmus`.

### Static Properties

| Property | Type | Description |
|---|---|---|
| `EmptyDocument` | `WmlDocument` | A minimal empty Word document used as a template |

### HtmlToWmlConverter Sample

```csharp
var settings = HtmlToWmlConverter.GetDefaultSettings();
settings.BaseUriForImages = "/images/";

var defaultCss = File.ReadAllText("default.css");
var authorCss = File.ReadAllText("styles.css");
var userCss = "";

var xhtml = XElement.Parse(@"
<html xmlns='http://www.w3.org/1999/xhtml'>
<head><title>Sample</title></head>
<body>
    <h1>Hello World</h1>
    <p>This is a <strong>sample</strong> document converted from HTML.</p>
    <table>
        <tr><td>Cell 1</td><td>Cell 2</td></tr>
        <tr><td>Cell 3</td><td>Cell 4</td></tr>
    </table>
</body>
</html>");

var doc = HtmlToWmlConverter.ConvertHtmlToWml(
    defaultCss, authorCss, userCss, xhtml, settings);
doc.SaveAs("output.docx");
```
