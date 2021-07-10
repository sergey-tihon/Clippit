---
uid: Tutorial.Word.DocumentBuilder.ISource
---
# Custom ISource implementation

Namespace: `Clippit.Word`

`ISource` abstractions allow to use `DocumentBuild` with custom content selectors.

```csharp
    public interface ISource : ICloneable
    {
        WmlDocument WmlDocument { get; set; }

        bool KeepSections { get; set; }
        public bool DiscardHeadersAndFootersInKeptSections { get; set; }

        string InsertId { get; set; }

        IEnumerable<XElement> GetElements(WordprocessingDocument document);
    }
```

### RecursiveTableCellSource

Allow to reference tables inside tables

```csharp
    [Serializable]
    public class TableCellReference
    {
        public int TableElementIndex { get; set; }

        public int RowIndex { get; set; }

        public int CellIndex { get; set; }
    }

    [Serializable]
    public class RecursiveTableCellSource : ISource
    {
        public WmlDocument WmlDocument
        {
            get => _wmlDocument;
            set => _wmlDocument = value;
        }

        [NonSerialized] private WmlDocument _wmlDocument;


        public bool KeepSections { get; set; }
        public bool DiscardHeadersAndFootersInKeptSections { get; set; }

        public string InsertId { get; set; }


        public List<TableCellReference> TableCellReferences { get; set; }

        public int Start { get; set; }

        public int Count { get; set; }


        public IEnumerable<XElement> GetElements(WordprocessingDocument document)
        {
            var body = document.MainDocumentPart.GetXDocument().Root?.Element(W.body);
            if (body is null)
            {
                throw new DocumentBuilderException(
                    "Unsupported document - contains no body element in the correct namespace");
            }

            var elements = body.Elements();
            foreach (var cellRef in TableCellReferences)
            {
                var table = elements.Skip(cellRef.TableElementIndex).FirstOrDefault();
                if (table is null || table.Name != W.tbl)
                {
                    throw new DocumentBuilderException(
                        $"Invalid {nameof(RecursiveTableCellSource)} - element {cellRef.TableElementIndex} is '{table?.Name}' but expected {W.tbl}");
                }

                var row = table.Elements(W.tr).Skip(cellRef.RowIndex).FirstOrDefault();
                if (row is null)
                {
                    throw new DocumentBuilderException(
                        $"Invalid {nameof(RecursiveTableCellSource)} - row {cellRef.RowIndex} does not exist");
                }

                var cell = row.Elements(W.tc).Skip(cellRef.CellIndex).FirstOrDefault();
                if (cell is null)
                {
                    throw new DocumentBuilderException(
                        $"Invalid {nameof(RecursiveTableCellSource)} - cell {cellRef.CellIndex} in the row {cellRef.RowIndex} does not exist");
                }

                elements = cell.Elements();
            }

            return elements
                .Skip(Start)
                .Take(Count)
                .ToList();
        }

        public object Clone() =>
            new RecursiveTableCellSource
            {
                WmlDocument = WmlDocument,
                KeepSections = KeepSections,
                DiscardHeadersAndFootersInKeptSections = DiscardHeadersAndFootersInKeptSections,
                InsertId = InsertId,
                TableCellReferences =
                    TableCellReferences.Select(x =>
                        new TableCellReference
                        {
                            TableElementIndex = x.TableElementIndex,
                            RowIndex = x.RowIndex,
                            CellIndex = x.CellIndex,
                        }).ToList(),
                Start = Start,
                Count = Count
            };
    }
```