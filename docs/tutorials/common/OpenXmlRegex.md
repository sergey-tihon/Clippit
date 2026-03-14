---
uid: Tutorial.Common.OpenXmlRegex
---
# OpenXml Regex

Namespace: `Clippit`

Search and replace text in Word and PowerPoint documents using regular expressions,
with optional revision tracking support.

```csharp
public class OpenXmlRegex {
    public static int Match(IEnumerable<XElement> content, Regex regex)
    {...}

    public static int Match(
        IEnumerable<XElement> content, Regex regex, Action<XElement, Match> found)
    {...}

    public static int Replace(
        IEnumerable<XElement> content, Regex regex, string replacement,
        Func<XElement, Match, bool> doReplacement)
    {...}

    public static int Replace(
        IEnumerable<XElement> content, Regex regex, string replacement,
        Func<XElement, Match, bool> doReplacement, bool coalesceContent)
    {...}

    public static int Replace(
        IEnumerable<XElement> content, Regex regex, string replacement,
        Func<XElement, Match, bool> doReplacement,
        bool trackRevisions, string author)
    {...}
}
```

`OpenXmlRegex` works on paragraph-level content elements (Word `w:p` and PowerPoint `a:p` elements).
It consolidates runs within each paragraph into a single logical string, applies the regex, and then
reconstructs the runs preserving original formatting.

#### Key Features

- **Works across runs** -- text split across multiple runs is matched as a single string
- **Preserves formatting** -- replaced text inherits the formatting of the first matched run
- **Revision tracking** -- replacements in Word content can generate tracked changes
  with a specified author (not supported for PowerPoint)
- **Callback control** -- the `doReplacement` callback controls whether each match is
  replaced, enabling selective replacement (e.g., replace only the first match)

### Match Sample

```csharp
using var doc = WordprocessingDocument.Open("input.docx", false);
var body = doc.MainDocumentPart.GetXDocument().Root.Element(W.body);
var paragraphs = body.Descendants(W.p);

var regex = new Regex(@"\b\d{3}-\d{2}-\d{4}\b"); // SSN pattern

// Count matches
var count = OpenXmlRegex.Match(paragraphs, regex);
Console.WriteLine($"Found {count} matches");

// Inspect each match
OpenXmlRegex.Match(paragraphs, regex, (element, match) =>
{
    Console.WriteLine($"Found: {match.Value}");
});
```

### Replace Sample

```csharp
var wmlDoc = new WmlDocument("input.docx");
using var streamDoc = new OpenXmlMemoryStreamDocument(wmlDoc);
using var doc = streamDoc.GetWordprocessingDocument();
var body = doc.MainDocumentPart.GetXDocument().Root.Element(W.body);
var paragraphs = body.Descendants(W.p).ToList();

var regex = new Regex(@"PLACEHOLDER");
var replacementCount = OpenXmlRegex.Replace(
    paragraphs, regex, "Actual Value", null);
doc.MainDocumentPart.PutXDocument();

Console.WriteLine($"Replaced {replacementCount} occurrences");
```

### Replace with Revision Tracking

```csharp
var regex = new Regex(@"old text");
OpenXmlRegex.Replace(
    paragraphs,
    regex,
    "new text",
    doReplacement: null,
    trackRevisions: true,
    author: "Auto-Replace Tool"
);
```
