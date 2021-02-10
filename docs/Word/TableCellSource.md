# Table Cell Source

Namespace: `Clippit.Word`

`TableCellSource` allow to reference content of table cells.

### Document Builder sample

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