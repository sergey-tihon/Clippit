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
all document parts (main document, headers, footers, footnotes, endnotes, comments, and styles).

#### Key Features

- **Accept revisions** -- applies all insertions and removes all deletions, producing a clean document
- **Reject revisions** -- reverses all insertions and restores all deletions, returning to the original state
- **Part-level control** -- accept revisions for individual parts or XML elements
- **Detection** -- check whether a document or part contains tracked revisions

When rejecting an inserted table cell, its grid columns are transferred to the preceding cell,
including multi-column spans. If that preceding cell has tracked property changes (`w:tcPrChange`),
its recorded properties are restored instead, without adding the inserted cell's columns.
Cell revisions in nested tables are processed independently and preserve enclosing cell content.
An inserted cell is still removed when its own properties also have tracked changes.
Cell-level content controls (`w:sdt`) and custom XML wrappers (`w:customXml`) are preserved while
their revised cells are removed and grid columns are transferred across wrapper boundaries.
The preceding-cell property-change exception also applies across these wrappers.

You can call the static API on `RevisionProcessor`:

```csharp
var cleanDoc = RevisionProcessor.AcceptRevisions(wmlDoc);
```

### AcceptRevisions Sample

```csharp
var wmlDoc = new WmlDocument("document_with_revisions.docx");

if (RevisionProcessor.HasTrackedRevisions(wmlDoc))
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
