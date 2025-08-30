// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
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
