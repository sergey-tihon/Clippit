---
uid: Tutorial.Word.DocumentBuilder.TableCellSource
---
# Table Cell Source

Namespace: `Clippit.Word`

`TableCellSource` allows extracting content from a specific table cell and including it
in the document being built.

```csharp
public class TableCellSource : ISource
{
    public TableCellSource() {...}
    public TableCellSource(WmlDocument source) {...}

    public WmlDocument WmlDocument { get; set; }
    public bool KeepSections { get; set; }
    public bool DiscardHeadersAndFootersInKeptSections { get; set; }
    public string InsertId { get; set; }

    // Table cell location
    public int TableElementIndex { get; set; }
    public int RowIndex { get; set; }
    public int CellIndex { get; set; }

    // Content range within the cell
    public int CellContentStart { get; set; }
    public int CellContentCount { get; set; }
}
```

| Property | Description |
|---|---|
| `TableElementIndex` | Zero-based index of the table element in the document body |
| `RowIndex` | Zero-based index of the row within the table |
| `CellIndex` | Zero-based index of the cell within the row |
| `CellContentStart` | Zero-based index of the first element to extract from the cell |
| `CellContentCount` | Number of elements to extract from the cell |

### DocumentBuilder Sample

```csharp {highlight:[12]}
    var document = new WmlDocument(sourceFilePath);
    
    var sources = new List<ISource>()
    {
        new Source(document)
        {
            // Select range of 5 elements (most frequently paragraphs)
            // starting from element with id 0
            Start = 0, 
            Count = 5
        },
        new TableCellSource(document)
        {
            // Reference the table (element with index 5)
            // Take 1st row and 3rd row in that row 
            TableElementIndex = 5, 
            RowIndex = 0, 
            CellIndex = 2, 
            // Select range of 2 elements inside the cell
            // starting from element with id 0
            CellContentStart = 0, 
            CellContentCount = 2
        }
    };

    var newDocument = DocumentBuilder.BuildDocument(sources);
    newDocument.SaveAs(destinationFilePath);
```