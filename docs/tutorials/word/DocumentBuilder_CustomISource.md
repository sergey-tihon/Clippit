---
uid: Tutorial.Word.DocumentBuilder.ISource
---
# Custom ISource Implementation

Namespace: `Clippit.Word`

The `ISource` interface allows using `DocumentBuilder` with custom content selectors.
Implement this interface to define your own logic for selecting which elements from a
source document should be included in the built document.

```csharp
public interface ISource : ICloneable
{
    WmlDocument WmlDocument { get; set; }
    bool KeepSections { get; set; }
    bool DiscardHeadersAndFootersInKeptSections { get; set; }
    string InsertId { get; set; }

    IEnumerable<XElement> GetElements(WordprocessingDocument document);
}
```

The `GetElements` method is called by `DocumentBuilder` to retrieve the content elements
from the source document. Your implementation controls which elements are returned.

## RecursiveTableCellSource

`RecursiveTableCellSource` is an example custom `ISource` implementation that allows
referencing content inside nested tables (tables within tables).

### TableCellReference

Each `TableCellReference` specifies one level of table nesting:

```csharp
public class TableCellReference
{
    public int TableElementIndex { get; set; }  // Index of table in current context
    public int RowIndex { get; set; }           // Row within the table
    public int CellIndex { get; set; }          // Cell within the row
}
```

### RecursiveTableCellSource Class

```csharp
public class RecursiveTableCellSource : ISource
{
    public WmlDocument WmlDocument { get; set; }
    public bool KeepSections { get; set; }
    public bool DiscardHeadersAndFootersInKeptSections { get; set; }
    public string InsertId { get; set; }

    // Navigation path through nested tables
    public List<TableCellReference> TableCellReferences { get; set; }

    // Content range within the final cell
    public int Start { get; set; }
    public int Count { get; set; }
}
```

### RecursiveTableCellSource Sample

```csharp
var document = new WmlDocument("nested-tables.docx");

// Navigate to: body -> table[0] -> row[1] -> cell[2] -> table[0] -> row[0] -> cell[1]
// Then extract 3 elements starting from element 0
var source = new RecursiveTableCellSource
{
    WmlDocument = document,
    TableCellReferences = new List<TableCellReference>
    {
        new() { TableElementIndex = 0, RowIndex = 1, CellIndex = 2 },  // Outer table
        new() { TableElementIndex = 0, RowIndex = 0, CellIndex = 1 }   // Inner table
    },
    Start = 0,
    Count = 3
};

var sources = new List<ISource> { source };
var result = DocumentBuilder.BuildDocument(sources);
result.SaveAs("extracted-content.docx");
```