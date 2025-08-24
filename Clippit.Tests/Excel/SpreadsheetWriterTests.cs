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
                    new() { Value = "Name", Bold = true },
                    new()
                    {
                        Value = "Age",
                        Bold = true,
                        HorizontalCellAlignment = HorizontalCellAlignment.Left,
                    },
                    new()
                    {
                        Value = "Rate",
                        Bold = true,
                        HorizontalCellAlignment = HorizontalCellAlignment.Left,
                    },
                ],
                Rows =
                [
                    new()
                    {
                        Cells =
                        [
                            new() { CellDataType = CellDataType.String, Value = "Eric" },
                            new() { CellDataType = CellDataType.Number, Value = 50 },
                            new()
                            {
                                CellDataType = CellDataType.Number,
                                Value = (decimal)45.00,
                                FormatCode = "0.00",
                            },
                        ],
                    },
                    new()
                    {
                        Cells =
                        [
                            new() { CellDataType = CellDataType.String, Value = "Bob" },
                            new() { CellDataType = CellDataType.Number, Value = 42 },
                            new()
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
        public void SaveWorkbookToFile()
        {
            var wb = GetSimpleWorkbookDfn();
            var fileName = Path.Combine(TempDir, "SW001-Simple.xlsx");
            using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            Validate(sDoc, s_spreadsheetExpectedErrors);
        }

        [Test]
        public void SaveWorkbookToStream()
        {
            var wb = GetSimpleWorkbookDfn();
            using var stream = new MemoryStream();
            wb.WriteTo(stream);
            stream.Position = 0;
            using var sDoc = SpreadsheetDocument.Open(stream, false);
            Validate(sDoc, s_spreadsheetExpectedErrors);
        }

        [Test]
        public void SaveWorkbookWithTwoSheets()
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
            using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            Validate(fileName);
        }

        [Test]
        public void SaveTablesWithDates()
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
            Validate(fileName);
        }

        [Test]
        public void SaveAllDataTypes()
        {
            var wb = new WorkbookDfn
            {
                Worksheets =
                [
                    new()
                    {
                        Name = "MyFirstSheet",
                        ColumnHeadings =
                        [
                            new() { Value = "DataType", Bold = true },
                            new()
                            {
                                Value = "Value",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Right,
                            },
                        ],
                        Rows =
                        [
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "Boolean" },
                                    new() { CellDataType = CellDataType.Boolean, Value = true },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "Boolean" },
                                    new() { CellDataType = CellDataType.Boolean, Value = false },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "String" },
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = HorizontalCellAlignment.Right,
                                    },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "int" },
                                    new() { CellDataType = CellDataType.Number, Value = 100 },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "int?" },
                                    new() { CellDataType = CellDataType.Number, Value = (int?)100 },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "int? (is null)" },
                                    new() { CellDataType = CellDataType.Number, Value = null },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "uint" },
                                    new() { CellDataType = CellDataType.Number, Value = (uint)101 },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "long" },
                                    new() { CellDataType = CellDataType.Number, Value = long.MaxValue },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "float" },
                                    new() { CellDataType = CellDataType.Number, Value = (float)123.45 },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "double" },
                                    new() { CellDataType = CellDataType.Number, Value = 123.45 },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new() { CellDataType = CellDataType.String, Value = "decimal" },
                                    new() { CellDataType = CellDataType.Number, Value = (decimal)123.45 },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new()
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9),
                                        FormatCode = "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                ],
                            },
                            new()
                            {
                                Cells =
                                [
                                    new()
                                    {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTimeOffset(new DateTime(2012, 1, 8), TimeSpan.Zero),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new()
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
            Validate(fileName);
        }

        [Test]
        public void AddWorksheetToWorkbook()
        {
            var wb = GetSimpleWorkbookDfn();
            var fileName = Path.Combine(TempDir, "AddWorksheetToWorkbook.xlsx");
            using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            using (var sDoc = SpreadsheetDocument.Open(fileName, true))
                SpreadsheetWriter.AddWorksheet(sDoc, GetSimpleWorksheetDfn("MySecondSheet", "MySecondTable"));
            Validate(fileName);
        }

        private void Validate(string fileName)
        {
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            Validate(sDoc, s_spreadsheetExpectedErrors);
        }

        private static readonly List<string> s_spreadsheetExpectedErrors =
        [
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        ];
    }
}
#endif
