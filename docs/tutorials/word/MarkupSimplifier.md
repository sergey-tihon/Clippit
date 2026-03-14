---
uid: Tutorial.Word.MarkupSimplifier
---
# Simplify Markup

Namespace: `Clippit.Word`

Strip unnecessary markup from Word documents to simplify the underlying XML structure.

```csharp
public static class MarkupSimplifier {
    public static WmlDocument SimplifyMarkup(
        WmlDocument doc, SimplifyMarkupSettings settings)
    {...}

    public static void SimplifyMarkup(
        WordprocessingDocument doc, SimplifyMarkupSettings settings)
    {...}
}
```

`MarkupSimplifier` removes various categories of markup that are often unnecessary for
document processing, comparison, or conversion. Each category is controlled by a flag
in `SimplifyMarkupSettings`.

An instance method is also available directly on `WmlDocument`:

```csharp
var simplified = wmlDoc.SimplifyMarkup(settings);
```

### SimplifyMarkupSettings

All fields are `bool` and default to `false`.

| Field | Description |
|---|---|
| `AcceptRevisions` | Accept all tracked revisions before simplifying |
| `NormalizeXml` | Normalize the XML structure |
| `RemoveBookmarks` | Remove bookmark start/end elements |
| `RemoveComments` | Remove comments and comment references |
| `RemoveContentControls` | Remove structured document tags (content controls) |
| `RemoveEndAndFootNotes` | Remove endnote and footnote references and content |
| `RemoveFieldCodes` | Remove field codes, keeping field results |
| `RemoveGoBackBookmark` | Remove the `_GoBack` bookmark |
| `RemoveHyperlinks` | Remove hyperlink wrappers |
| `RemoveLastRenderedPageBreak` | Remove `lastRenderedPageBreak` elements |
| `RemoveMarkupForDocumentComparison` | Remove markup that interferes with document comparison (implies `RemoveRsidInfo`) |
| `RemovePermissions` | Remove permission start/end elements |
| `RemoveProof` | Remove proofing markup (spell check, grammar) |
| `RemoveRsidInfo` | Remove revision save ID attributes |
| `RemoveSmartTags` | Remove smart tag elements |
| `RemoveSoftHyphens` | Remove soft hyphen characters |
| `RemoveWebHidden` | Remove web-hidden paragraph marks |
| `ReplaceTabsWithSpaces` | Replace tab characters with spaces |

### Additional Methods

| Method | Description |
|---|---|
| `MergeAdjacentSuperfluousRuns(XElement)` | Merge adjacent runs with identical formatting |
| `TransformElementToSingleCharacterRuns(XElement)` | Split runs so each contains a single character |
| `TransformPartToSingleCharacterRuns(OpenXmlPart)` | Apply single-character run transform to a part |
| `TransformToSingleCharacterRuns(WordprocessingDocument)` | Apply single-character run transform to entire document |

### SimplifyMarkup Sample

```csharp
var wmlDoc = new WmlDocument("input.docx");

var settings = new SimplifyMarkupSettings
{
    RemoveComments = true,
    RemoveRsidInfo = true,
    RemoveProof = true,
    RemoveBookmarks = true,
    RemoveGoBackBookmark = true,
    RemoveSoftHyphens = true,
    RemoveLastRenderedPageBreak = true,
    RemoveContentControls = true,
    RemoveSmartTags = true
};

var simplified = wmlDoc.SimplifyMarkup(settings);
simplified.SaveAs("simplified.docx");
```

### Prepare for Comparison Sample

```csharp
var settings = new SimplifyMarkupSettings
{
    RemoveMarkupForDocumentComparison = true,
    AcceptRevisions = true
};

var doc1 = new WmlDocument("doc1.docx").SimplifyMarkup(settings);
var doc2 = new WmlDocument("doc2.docx").SimplifyMarkup(settings);

// Documents are now ready for structural comparison
```
