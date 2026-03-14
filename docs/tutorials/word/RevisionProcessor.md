---
uid: Tutorial.Word.RevisionProcessor
---
# Process Revisions

Namespace: `Clippit.Word`

Accept or reject tracked revisions in Word documents programmatically.

```csharp
public class RevisionProcessor {
    public static WmlDocument AcceptRevisions(WmlDocument document)
    {...}
    public static void AcceptRevisions(WordprocessingDocument doc)
    {...}

    public static WmlDocument RejectRevisions(WmlDocument document)
    {...}
    public static void RejectRevisions(WordprocessingDocument doc)
    {...}

    public static void AcceptRevisionsForPart(OpenXmlPart part)
    {...}
    public static XElement AcceptRevisionsForElement(XElement element)
    {...}

    public static bool HasTrackedRevisions(WmlDocument document)
    {...}
    public static bool HasTrackedRevisions(WordprocessingDocument doc)
    {...}
    public static bool PartHasTrackedRevisions(OpenXmlPart part)
    {...}
}
```

`RevisionProcessor` handles both accepting and rejecting tracked revisions across
all document parts (main document, headers, footers, footnotes, endnotes, and styles).

#### Key Features

- **Accept revisions** -- applies all insertions and removes all deletions, producing a clean document
- **Reject revisions** -- reverses all insertions and restores all deletions, returning to the original state
- **Part-level control** -- accept revisions for individual parts or XML elements
- **Detection** -- check whether a document or part contains tracked revisions

An instance method is also available directly on `WmlDocument`:

```csharp
var cleanDoc = wmlDoc.AcceptRevisions();
```

### AcceptRevisions Sample

```csharp
var wmlDoc = new WmlDocument("document_with_revisions.docx");

if (wmlDoc.HasTrackedRevisions())
{
    var accepted = RevisionProcessor.AcceptRevisions(wmlDoc);
    accepted.SaveAs("document_clean.docx");
}
```

### RejectRevisions Sample

```csharp
var wmlDoc = new WmlDocument("document_with_revisions.docx");

var rejected = RevisionProcessor.RejectRevisions(wmlDoc);
rejected.SaveAs("document_original.docx");
```

### AcceptRevisionsForPart Sample

```csharp
using var doc = WordprocessingDocument.Open("input.docx", true);

// Accept revisions only in the main document part
RevisionProcessor.AcceptRevisionsForPart(doc.MainDocumentPart);

doc.Save();
```
