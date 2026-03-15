---
uid: Tutorial.Common.MetricsGetter
---
# Document Metrics

Namespace: `Clippit`

Analyze Word, Excel, and PowerPoint documents and return detailed metrics as XML.

```csharp
public class MetricsGetter {
    public static XElement GetMetrics(string fileName, MetricsGetterSettings settings)
    {...}

    public static XElement GetDocxMetrics(WmlDocument wmlDoc, MetricsGetterSettings settings)
    {...}

    public static XElement GetXlsxMetrics(SmlDocument smlDoc, MetricsGetterSettings settings)
    {...}

    public static XElement GetPptxMetrics(PmlDocument pmlDoc, MetricsGetterSettings settings)
    {...}
}
```

`MetricsGetter` inspects Office documents and produces an `XElement` containing metrics such as
paragraph counts, character counts (by script), run counts, content control details, embedded
objects, hyperlinks, revision tracking status, namespaces, and validation errors.

The `GetMetrics` method auto-detects the file type by extension and delegates to
`GetDocxMetrics`, `GetXlsxMetrics`, or `GetPptxMetrics`.

### MetricsGetterSettings

| Property | Type | Default | Description |
|---|---|---|---|
| `IncludeTextInContentControls` | `bool` | `false` | Include text content of content controls in output |
| `IncludeXlsxTableCellData` | `bool` | `false` | Include cell data from Excel tables |
| `RetrieveNamespaceList` | `bool` | `false` | Include list of namespaces used in the document |
| `RetrieveContentTypeList` | `bool` | `false` | Include list of content types of document parts |

### GetMetrics Sample

```csharp
var settings = new MetricsGetterSettings
{
    RetrieveNamespaceList = true,
    RetrieveContentTypeList = true
};

var metrics = MetricsGetter.GetMetrics("document.docx", settings);
Console.WriteLine(metrics.ToString());
```

Output includes elements such as:
- `ParagraphCount`, `RunCount`, `AsciiCharCount`, `EastAsiaCharCount`
- `RevisionTracking` (whether the document has tracked revisions)
- `ContentControls` (if present)
- `Namespaces` (if `RetrieveNamespaceList` is `true`)
- Validation errors from the OpenXml SDK validator

### GetDocxMetrics Sample

```csharp
var wmlDoc = new WmlDocument("report.docx");
var settings = new MetricsGetterSettings
{
    IncludeTextInContentControls = true
};

var metrics = MetricsGetter.GetDocxMetrics(wmlDoc, settings);

// Extract specific values
var paragraphCount = (int?)metrics.Element("ParagraphCount");
var hasRevisions = (bool?)metrics.Element("RevisionTracking")?.Attribute("Val");
Console.WriteLine($"Paragraphs: {paragraphCount}, Has revisions: {hasRevisions}");
```

### GetXlsxMetrics Sample

```csharp
var smlDoc = new SmlDocument("data.xlsx");
var settings = new MetricsGetterSettings
{
    IncludeXlsxTableCellData = true
};

var metrics = MetricsGetter.GetXlsxMetrics(smlDoc, settings);
Console.WriteLine(metrics.ToString());
```
