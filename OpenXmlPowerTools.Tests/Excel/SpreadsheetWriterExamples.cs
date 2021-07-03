using System;
using System.IO;
using Clippit.Excel;
using Xunit;
using Xunit.Abstractions;

namespace Clippit.Tests.Excel
{
    public class SpreadsheetWriterExamples : TestsBase
    {
        protected SpreadsheetWriterExamples(ITestOutputHelper log) : base(log)
        {
        }
        
        [Fact]
        public void Example01()
        {
            var wb = new WorkbookDfn
            {
                Worksheets = new[]
                {
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        TableName = "NamesAndRates",
                        ColumnHeadings = new[]
                        {
                            new CellDfn
                            {
                                Value = "Name",
                                Bold = true,
                            },
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
                            }
                        },
                        Rows = new[]
                        {
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Eric",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 50,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)45.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Bob",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = 42,
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)78.00,
                                        FormatCode = "0.00",
                                    },
                                }
                            },
                        }
                    }
                }
            };
            
            var fileName = Path.Combine(TempDir, "Sw_Example1.xlsx");
            using var stream = File.Open(fileName, FileMode.OpenOrCreate);
            wb.WriteTo(stream);
        }
        
        [Fact]
        public void Example02()
        {
            var wb = new WorkbookDfn
            {
                Worksheets = new[]
                {
                    new WorksheetDfn
                    {
                        Name = "MyFirstSheet",
                        ColumnHeadings = new[]
                        {
                            new CellDfn
                            {
                                Value = "DataType",
                                Bold = true,
                            },
                            new CellDfn
                            {
                                Value = "Value",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Right,
                            },
                        },
                        Rows = new[]
                        {
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Boolean,
                                        Value = true,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Boolean,
                                        Value = false,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "String",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = HorizontalCellAlignment.Right,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "int",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (int)100,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "int?",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (int?)100,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "int? (is null)",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = null,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "uint",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (uint)101,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "long",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = Int64.MaxValue,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "float",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (float)123.45,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "double",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (double)123.45,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.String,
                                        Value = "decimal",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)123.45,
                                    },
                                }
                            },
                            new RowDfn
                            {
                                Cells = new[]
                                {
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 8),
                                        FormatCode = "mm-dd-yy",
                                    },
                                    new CellDfn {
                                        CellDataType = CellDataType.Date,
                                        Value = new DateTime(2012, 1, 9),
                                        FormatCode = "mm-dd-yy",
                                        Bold = true,
                                        HorizontalCellAlignment = HorizontalCellAlignment.Center,
                                    },
                                }
                            },
                        }
                    }
                }
            };
            
            var fileName = Path.Combine(TempDir, "Sw_Example2.xlsx");
            using var stream = File.Open(fileName, FileMode.OpenOrCreate);
            wb.WriteTo(stream);
        }
    }
}
