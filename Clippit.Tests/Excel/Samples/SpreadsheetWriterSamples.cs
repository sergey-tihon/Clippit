using Clippit.Excel;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Excel.Samples
{
    public class SpreadsheetWriterSamples(ITestOutputHelper log) : TestsBase(log)
    {
        [Fact]
        public void Sample1()
        {
            var wb = new WorkbookDfn
            {
                Worksheets =
                [
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        TableName = "NamesAndRates",
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
                    },
                ],
            };

            var fileName = Path.Combine(TempDir, "Sw_Example1.xlsx");
            using var stream = File.Open(fileName, FileMode.OpenOrCreate);
            wb.WriteTo(stream);
        }

        [Fact]
        public void Sample2()
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
                        ],
                    },
                ],
            };

            var fileName = Path.Combine(TempDir, "Sw_Example2.xlsx");
            using var stream = File.Open(fileName, FileMode.OpenOrCreate);
            wb.WriteTo(stream);
        }

        [Fact]
        public void CanEncodeInvalidXmlCharacters()
        {
            var wb = new WorkbookDfn
            {
                Worksheets =
                [
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        Rows =
                        [
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "Invalid character: \uFFFF",
                                    },
                                ],
                            },
                        ],
                    },
                ],
                Options = new WorkbookDfnOptions { InvalidCharterBehavior = InvalidCharterBehavior.Remove },
            };

            var fileName = Path.Combine(TempDir, $"{nameof(CanEncodeInvalidXmlCharacters)}.xlsx");
            using var stream = File.Open(fileName, FileMode.OpenOrCreate);
            wb.WriteTo(stream);
        }

        // write another test but this time throw an exception
        [Fact]
        public void CanThrowExceptionOnInvalidXmlCharacters()
        {
            var wb = new WorkbookDfn
            {
                Worksheets =
                [
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        Rows =
                        [
                            new RowDfn
                            {
                                Cells =
                                [
                                    new CellDfn
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "Invalid character: \uFFFF",
                                    },
                                ],
                            },
                        ],
                    },
                ],
                Options = new WorkbookDfnOptions { InvalidCharterBehavior = InvalidCharterBehavior.ThrowException },
            };

            var fileName = Path.Combine(TempDir, $"{nameof(CanThrowExceptionOnInvalidXmlCharacters)}.xlsx");
            using var stream = File.Open(fileName, FileMode.OpenOrCreate);

            var exception = Assert.Throws<ArgumentException>(() => wb.WriteTo(stream));
            Assert.Contains("invalid character", exception.Message, StringComparison.InvariantCultureIgnoreCase);
        }
    }
}
