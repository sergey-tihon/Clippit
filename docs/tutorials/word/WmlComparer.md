---
uid: Tutorial.Word.WmlComparer
---
# Compare Documents

Namespace: `Clippit`

Compare two Word documents and produce a result document with tracked revisions showing the differences.
Consolidate multiple revised documents against an original, or extract a list of revisions from a comparison result.

```csharp
public static class WmlComparer {
    public static WmlDocument Compare(
        WmlDocument source1, WmlDocument source2, WmlComparerSettings settings)
    {...}

    public static WmlDocument Consolidate(
        WmlDocument original,
        List<WmlRevisedDocumentInfo> revisedDocumentInfoList,
        WmlComparerSettings settings)
    {...}

    public static List<WmlComparerRevision> GetRevisions(
        WmlDocument source, WmlComparerSettings settings)
    {...}
}
```

### Compare

Compares two Word documents and returns a new document containing tracked revisions (insertions and deletions)
that represent the differences between `source1` and `source2`. The comparison operates at the word level
by default, using `WordSeparators` from `WmlComparerSettings` to split text into tokens.

### Consolidate

Merges multiple revised versions of a document against a common original. Each revised document is provided
with a revisor name and color via `WmlRevisedDocumentInfo`. The result contains tracked revisions from all
revisors. By default, revisions are wrapped in a comparison table (controlled by `WmlComparerConsolidateSettings.ConsolidateWithTable`).

### GetRevisions

Extracts a flat list of `WmlComparerRevision` objects from a document that contains tracked revisions
(typically the output of `Compare` or `Consolidate`). Each revision includes the revision type
(`Inserted` or `Deleted`), text content, author, and date.

### WmlComparerSettings

| Property | Type | Default |
|---|---|---|
| `WordSeparators` | `char[]` | `[' ', '-', ')', '(', ';', ',']` |
| `AuthorForRevisions` | `string` | `"Open-Xml-PowerTools"` |
| `DateTimeForRevisions` | `string` | `DateTime.Now.ToString("o")` |
| `DetailThreshold` | `double` | `0.15` |
| `CaseInsensitive` | `bool` | `false` |
| `CultureInfo` | `CultureInfo` | `null` |
| `LogCallback` | `Action<string>` | `null` |
| `StartingIdForFootnotesEndnotes` | `int` | `1` |
| `DebugTempFileDi` | `DirectoryInfo` | `null` |

### Compare Sample

```csharp
var source1 = new WmlDocument("Original.docx");
var source2 = new WmlDocument("Revised.docx");

var settings = new WmlComparerSettings
{
    AuthorForRevisions = "Comparison Tool",
    DetailThreshold = 0.15
};

var comparedDoc = WmlComparer.Compare(source1, source2, settings);
comparedDoc.SaveAs("Comparison.docx");
```

### Consolidate Sample

```csharp
var original = new WmlDocument("Original.docx");

var revisedDocs = new List<WmlRevisedDocumentInfo>
{
    new()
    {
        RevisedDocument = new WmlDocument("Revised_Alice.docx"),
        Revisor = "Alice",
        Color = Color.LightBlue
    },
    new()
    {
        RevisedDocument = new WmlDocument("Revised_Bob.docx"),
        Revisor = "Bob",
        Color = Color.LightGreen
    }
};

var settings = new WmlComparerSettings();

var consolidated = WmlComparer.Consolidate(original, revisedDocs, settings);
consolidated.SaveAs("Consolidated.docx");
```

### GetRevisions Sample

```csharp
var comparedDoc = WmlComparer.Compare(source1, source2, settings);

var revisions = WmlComparer.GetRevisions(comparedDoc, settings);
foreach (var rev in revisions)
{
    Console.WriteLine($"{rev.RevisionType}: \"{rev.Text}\" by {rev.Author} on {rev.Date}");
}
```
