// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
using Clippit;
using Clippit.Excel;
using DocumentFormat.OpenXml.Packaging;

#if !ELIDE_XUNIT_TESTS
namespace Clippit.Tests.Excel
{
    public class SpreadsheetWriterTests : TestsBase
    {
        private static WorkbookDfn GetSimpleWorkbookDfn() =>
            new() { Worksheets = [GetSimpleWorksheetDfn("MyFirstSheet", "NamesAndRates")] };

        private static WorksheetDfn GetSimpleWorksheetDfn(string name, string table) =>
            new()
            {
                Name = name,
                TableName = table,
                ColumnHeadings =
                [
                    new CellDfn { Value = "Name", Bold = true },
                    new CellDfn
                    {
                        Value = "Age",
                        Bold = true,
                        HorizontalCellAlignment = HorizontalCellAlignment.Left,
                    },
                    new CellDfn
                    {
                        Value = "Rate",
                        Bold = true,
                        HorizontalCellAlignment = HorizontalCellAlignment.Left,
                    },
                ],
                Rows =
                [
                    new RowDfn
                    {
                        Cells =
                        [
                            new CellDfn { CellDataType = CellDataType.String, Value = "Eric" },
                            new CellDfn { CellDataType = CellDataType.Number, Value = 50 },
                            new CellDfn
                            {
                                CellDataType = CellDataType.Number,
                                Value = (decimal)45.00,
                                FormatCode = "0.00",
                            },
                        ],
                    },
                    new RowDfn
                    {
                        Cells =
                        [
                            new CellDfn { CellDataType = CellDataType.String, Value = "Bob" },
                            new CellDfn { CellDataType = CellDataType.Number, Value = 42 },
                            new CellDfn
                            {
                                CellDataType = CellDataType.Number,
                                Value = (decimal)78.00,
                                FormatCode = "0.00",
                            },
                        ],
                    },
                ],
            };

        [Test]
        public async Task SaveWorkbookToFile()
        {
            var wb = GetSimpleWorkbookDfn();
            var fileName = Path.Combine(TempDir, "SW001-Simple.xlsx");
            await using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            await Validate(sDoc, s_spreadsheetExpectedErrors).ConfigureAwait(false);
        }

        [Test]
        public async Task SaveWorkbookToStream()
        {
            var wb = GetSimpleWorkbookDfn();
            using var stream = new MemoryStream();
            wb.WriteTo(stream);
            stream.Position = 0;
            using var sDoc = SpreadsheetDocument.Open(stream, false);
            await Validate(sDoc, s_spreadsheetExpectedErrors).ConfigureAwait(false);
        }

        [Test]
        public async Task SaveWorkbookWithTwoSheets()
        {
            var wb = new WorkbookDfn
            {
                Worksheets =
                [
                    GetSimpleWorksheetDfn("MyFirstSheet", "NamesAndRates1"),
                    GetSimpleWorksheetDfn("MySecondSheet", "NamesAndRates2"),
                ],
            };
            var fileName = Path.Combine(TempDir, "SW001_TwoSheets.xlsx");
            await using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            await Validate(fileName).ConfigureAwait(false);
        }

        [Test]
        public async Task SaveTablesWithDates()
        {
            WorksheetDfn GetSheet(string name, string tableName) =>
                new()
                {
                    Name = name,
                    TableName = tableName,
                    ColumnHeadings =
                    [
                        new CellDfn
                        {
                            CellDataType = CellDataType.String,
                            Bold = true,
                            Value = "Date",
                        },
                    ],
                    Rows =
                    [
                        new RowDfn { Cells = [null] },
                        new RowDfn
                        {
                            Cells =
                            [
                                new CellDfn
                                {
                                    CellDataType = CellDataType.Date,
                                    Value = null,
                                    FormatCode = "mm-dd-yy",
                                },
                            ],
                        },
                        new RowDfn
                        {
                            Cells =
                            [
                                new CellDfn
                                {
                                    CellDataType = CellDataType.Date,
                                    Value = DateTime.Now,
                                    FormatCode = "mm-dd-yy",
                                },
                            ],
                        },
                    ],
                };
            var wb = new WorkbookDfn { Worksheets = [GetSheet("Sheet1", "Table1"), GetSheet("Sheet2", "Table2")] };
            var fileName = Path.Combine(TempDir, "SW001_TableWithDates.xlsx");
            using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            await Validate(fileName).ConfigureAwait(false);
        }

        [Test]
        public async Task SaveAllDataTypes()
        {
            var wb = new WorkbookDfn
            {
                Worksheets =
                [
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        ColumnHeadings =
                        [
                            new CellDfn { Value = "DataType", Bold = true },
                            new CellDfn
                            {
                                Value = "Value",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Right,
                            },
                        ],
                        Rows =
                        [
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "Boolean" },
                                    new CellDfn { CellDataType = CellDataType.Boolean, Value = true },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "Boolean" },
                                    new CellDfn { CellDataType = CellDataType.Boolean, Value = false },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "String" },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = HorizontalCellAlignment.Right,
                                    },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "int" },
                                    new CellDfn { CellDataType = CellDataType.Number, Value = 100 },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "int?" },
                                    new CellDfn { CellDataType = CellDataType.Number, Value = (int?)100 },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "int? (is null)" },
                                    new CellDfn { CellDataType = CellDataType.Number, Value = null },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "uint" },
                                    new CellDfn { CellDataType = CellDataType.Number, Value = (uint)101 },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "long" },
                                    new CellDfn { CellDataType = CellDataType.Number, Value = long.MaxValue },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "float" },
                                    new CellDfn { CellDataType = CellDataType.Number, Value = (float)123.45 },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "double" },
                                    new CellDfn { CellDataType = CellDataType.Number, Value = 123.45 },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn { CellDataType = CellDataType.String, Value = "decimal" },
                                    new CellDfn { CellDataType = CellDataType.Number, Value = (decimal)123.45 },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9),
                                        FormatCode = "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                ],
                            },
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTimeOffset(new DateTime(2012, 1, 8), TimeSpan.Zero),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTimeOffset(new DateTime(2012, 1, 9), TimeSpan.Zero),
                                        FormatCode = "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                ],
                            },
                        ],
                    },
                ],
            };
            var fileName = Path.Combine(TempDir, "SW002-DataTypes.xlsx");
            using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            await Validate(fileName).ConfigureAwait(false);
        }

        // Verifies that table definitions are NOT created when TableName is set but no data rows exist.
        // XLSX spec requires tables to have at least one data row. Creating a table with only headers
        // produces a malformed file that fails to open in Excel.
        [Test]
        public async Task SaveWorksheetWithTableNameButNoDataRows()
        {
            var wb = new WorkbookDfn
            {
                Worksheets =
                [
                    new WorksheetDfn
                    {
                        Name = "EmptySheet",
                        TableName = "EmptyTable",
                        ColumnHeadings =
                        [
                            new CellDfn { Value = "Name", Bold = true },
                            new CellDfn { Value = "Age", Bold = true },
                        ],
                        Rows = [],
                    },
                ],
            };
            var fileName = Path.Combine(TempDir, "SW003-EmptyTable.xlsx");
            await using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);

            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            var worksheetPart = sDoc.WorkbookPart.WorksheetParts.First();
            var tableDefParts = worksheetPart.GetPartsOfType<TableDefinitionPart>();

            await Assert.That(tableDefParts).IsEmpty();

            await Validate(sDoc, s_spreadsheetExpectedErrors).ConfigureAwait(false);
        }

        [Test]
        public async Task AddWorksheetToWorkbook()
        {
            var wb = GetSimpleWorkbookDfn();
            var fileName = Path.Combine(TempDir, "AddWorksheetToWorkbook.xlsx");
            await using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            using (var sDoc = SpreadsheetDocument.Open(fileName, true))
                SpreadsheetWriter.AddWorksheet(sDoc, GetSimpleWorksheetDfn("MySecondSheet", "MySecondTable"));
            await Validate(fileName).ConfigureAwait(false);
        }

        // Verifies that numFmts count attribute stays in sync when multiple distinct custom
        // format codes are registered across worksheets (previously the count was not incremented
        // after the first custom numFmt was added, leaving stale count="1" regardless of how
        // many custom formats had been registered).
        [Test]
        public async Task SW004_MultipleCustomFormatCodes_NumFmtsCountIsCorrect()
        {
            var wb = new WorkbookDfn
            {
                Worksheets =
                [
                    new WorksheetDfn
                    {
                        Name = "StringSheet",
                        Rows =
                        [
                            new RowDfn
                            {
                                Cells = [new CellDfn { CellDataType = CellDataType.String, Value = "Hello" }],
                            },
                        ],
                    },
                    new WorksheetDfn
                    {
                        Name = "FormattedSheet",
                        Rows =
                        [
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = 1234.5,
                                        FormatCode = "#,##0.000",
                                    },
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = 0.75,
                                        FormatCode = "\"Rate:\" 0.00%",
                                    },
                                ],
                            },
                        ],
                    },
                ],
            };

            var fileName = Path.Combine(TempDir, "SW004-MultipleCustomFormats.xlsx");
            await using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);

            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            var stylesXDoc = sDoc.WorkbookPart.WorkbookStylesPart.GetXDocument();

            var numFmtsEl = stylesXDoc.Root.Element(S.numFmts);
            await Assert.That(numFmtsEl).IsNotNull();

            var declaredCount = (int)numFmtsEl.Attribute("count");
            var actualCount = numFmtsEl.Elements().Count();
            await Assert.That(declaredCount).IsEqualTo(actualCount);

            await Validate(sDoc, s_spreadsheetExpectedErrors).ConfigureAwait(false);
        }

        // Reproduces issue #64: string sheet first, date sheet second must produce a valid workbook.
        // Also verifies that custom numFmt IDs are >= 164 (ECMA-376 §18.8.30 reserves 0–163 for
        // built-in formats) and that repeating the same FormatCode does not allocate a duplicate entry.
        [Test]
        public async Task SW005_StringSheetFirst_DateSheetSecond_CustomFormatId()
        {
            static WorksheetDfn MakeStringSheet(string name) =>
                new()
                {
                    Name = name,
                    ColumnHeadings =
                    [
                        new CellDfn { Value = "col1_string", Bold = true },
                        new CellDfn { Value = "col2_string", Bold = true },
                    ],
                    Rows =
                    [
                        new RowDfn
                        {
                            Cells =
                            [
                                new CellDfn { CellDataType = CellDataType.String, Value = "Hello" },
                                new CellDfn { CellDataType = CellDataType.String, Value = "world" },
                            ],
                        },
                        new RowDfn
                        {
                            Cells =
                            [
                                new CellDfn
                                {
                                    CellDataType = CellDataType.Date,
                                    Value = new DateTime(2023, 4, 20),
                                    FormatCode = "dd-MM-yyyy",
                                },
                                new CellDfn { CellDataType = CellDataType.String, Value = "another" },
                            ],
                        },
                    ],
                };

            static WorksheetDfn MakeDateSheet(string name) =>
                new()
                {
                    Name = name,
                    ColumnHeadings =
                    [
                        new CellDfn { Value = "col1_date", Bold = true },
                        new CellDfn { Value = "col2_int", Bold = true },
                    ],
                    Rows =
                    [
                        new RowDfn
                        {
                            Cells =
                            [
                                new CellDfn
                                {
                                    CellDataType = CellDataType.Date,
                                    Value = new DateTime(2023, 4, 19),
                                    FormatCode = "dd-MM-yyyy",
                                },
                                new CellDfn { CellDataType = CellDataType.Number, Value = -1 },
                            ],
                        },
                    ],
                };

            // String-first, date-second (the failing order from issue #64)
            var wb = new WorkbookDfn
            {
                Worksheets = [MakeStringSheet("StringSheet"), MakeDateSheet("DateSheet")],
            };
            var fileName = Path.Combine(TempDir, "SW005-StringFirst-DateSecond.xlsx");
            await using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);

            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            var stylesXDoc = sDoc.WorkbookPart.WorkbookStylesPart.GetXDocument();

            var numFmtsEl = stylesXDoc.Root.Element(S.numFmts);
            await Assert.That(numFmtsEl).IsNotNull();

            // All custom numFmt IDs must be in the user-defined range (>= 164).
            var customIds = numFmtsEl.Elements(S.numFmt).Select(e => (int)e.Attribute("numFmtId")).ToList();
            foreach (var id in customIds)
                await Assert.That(id).IsGreaterThanOrEqualTo(164);

            // The same FormatCode used on both sheets must not produce duplicate numFmt entries.
            var formatCodes = numFmtsEl
                .Elements(S.numFmt)
                .Select(e => (string)e.Attribute("formatCode"))
                .ToList();
            await Assert.That(formatCodes.Distinct().Count()).IsEqualTo(formatCodes.Count);

            // Specifically, there should be exactly one numFmt entry for the custom format code used ("dd-MM-yyyy").
            const string targetFormatCode = "dd-MM-yyyy";
            var targetFormatCodeCount = formatCodes.Count(fc => fc == targetFormatCode);
            await Assert.That(targetFormatCodeCount).IsEqualTo(1);

            // count attribute must match actual element count, and remain 1 for this workbook.
            var declaredCount = (int)numFmtsEl.Attribute("count");
            await Assert.That(declaredCount).IsEqualTo(numFmtsEl.Elements().Count());
            await Assert.That(declaredCount).IsEqualTo(1);

            await Validate(sDoc, s_spreadsheetExpectedErrors).ConfigureAwait(false);
        }

        private async Task Validate(string fileName)
        {
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            await Validate(sDoc, s_spreadsheetExpectedErrors).ConfigureAwait(false);
        }

        private static readonly List<string> s_spreadsheetExpectedErrors =
        [
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        ];
    }
}
#endif
