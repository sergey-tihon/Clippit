// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Clippit.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using Sw = Clippit;
using Xunit;
using Xunit.Abstractions;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class SwTests
    {
        private readonly ITestOutputHelper _log;

        public SwTests(ITestOutputHelper log)
        {
            this._log = log;
        }
        
        private static WorkbookDfn GetSimpleWorkbookDfn() => new()
        {
            Worksheets = new[] { GetSimpleWorksheetDfn("MyFirstSheet","NamesAndRates") }
        };

        private static WorksheetDfn GetSimpleWorksheetDfn(string name, string table) => new()
        {
            Name = name,
            TableName = table,
            ColumnHeadings =
                new CellDfn[]
                {
                    new() { Value = "Name", Bold = true, },
                    new() { Value = "Age", Bold = true, HorizontalCellAlignment = HorizontalCellAlignment.Left, },
                    new() { Value = "Rate", Bold = true, HorizontalCellAlignment = HorizontalCellAlignment.Left, }
                },
            Rows = new RowDfn[]
            {
                new()
                {
                    Cells = new CellDfn[]
                    {
                        new() { CellDataType = CellDataType.String, Value = "Eric", },
                        new() { CellDataType = CellDataType.Number, Value = 50, },
                        new()
                        {
                            CellDataType = CellDataType.Number,
                            Value = (decimal)45.00,
                            FormatCode = "0.00",
                        },
                    }
                },
                new()
                {
                    Cells = new CellDfn[]
                    {
                        new() { CellDataType = CellDataType.String, Value = "Bob", },
                        new() { CellDataType = CellDataType.Number, Value = 42, },
                        new()
                        {
                            CellDataType = CellDataType.Number,
                            Value = (decimal)78.00,
                            FormatCode = "0.00",
                        },
                    }
                },
            }
        };
            
        
        [Fact]
        public void SW001_Simple()
        {
            var wb = GetSimpleWorkbookDfn();
            
            var fileName = Path.Combine(Sw.TestUtil.TempDir.FullName, "SW001-Simple.xlsx");
            using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            Validate(sDoc);
        }
        
        [Fact]
        public void SW001_SimpleToStream()
        {
            var wb = GetSimpleWorkbookDfn();

            using var stream = new MemoryStream();
            wb.WriteTo(stream);
            stream.Position = 0;

            using var sDoc = SpreadsheetDocument.Open(stream, false);
            Validate(sDoc);
        }
        
        [Fact]
        public void SW001_TwoSheets()
        {
            var wb = new WorkbookDfn
            {
                Worksheets = new[]
                {
                    GetSimpleWorksheetDfn("MyFirstSheet","NamesAndRates1"),
                    GetSimpleWorksheetDfn("MySecondSheet","NamesAndRates2")
                }
            };

            var fileName = Path.Combine(Sw.TestUtil.TempDir.FullName, "SW001_TwoSheets.xlsx");
            using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);

            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            Validate(sDoc);
        }

        [Fact]
        public void SW002_AllDataTypes()
        {
            var wb = new WorkbookDfn
            {
                Worksheets = new WorksheetDfn[]
                {
                    new()
                    {
                        Name = "MyFirstSheet",
                        ColumnHeadings = new CellDfn[]
                        {
                            new()
                            {
                                Value = "DataType",
                                Bold = true,
                            },
                            new()
                            {
                                Value = "Value",
                                Bold = true,
                                HorizontalCellAlignment = HorizontalCellAlignment.Right,
                            },
                        },
                        Rows = new RowDfn[]
                        {
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Boolean,
                                        Value = true,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "Boolean",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Boolean,
                                        Value = false,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "String",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "A String",
                                        HorizontalCellAlignment = HorizontalCellAlignment.Right,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "int",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (int)100,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "int?",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (int?)100,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "int? (is null)",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = null,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "uint",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (uint)101,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "long",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = Int64.MaxValue,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "float",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (float)123.45,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "double",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (double)123.45,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
                                    new()
                                    {
                                        CellDataType = CellDataType.String,
                                        Value = "decimal",
                                    },
                                    new()
                                    {
                                        CellDataType = CellDataType.Number,
                                        Value = (decimal)123.45,
                                    },
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
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
                                }
                            },
                            new()
                            {
                                Cells = new CellDfn[]
                                {
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
                                }
                            },
                        }
                    }
                }
            };

            var fileName = Path.Combine(Sw.TestUtil.TempDir.FullName, "SW002-DataTypes.xlsx");
            using (var stream = File.Open(fileName, FileMode.OpenOrCreate))
                wb.WriteTo(stream);
            
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            Validate(sDoc);
        }

        private void Validate(SpreadsheetDocument sDoc)
        {
            var v = new OpenXmlValidator();
            var errors = v.Validate(sDoc)
                .Where(ve => !s_expectedErrors.Contains(ve.Description))
                .ToList();

            // if a test fails validation post-processing, then can use this code to determine the SDK
            // validation error(s).
            foreach (var item in errors)
            {
                _log.WriteLine(item.Description);
            }
            
            Assert.Empty(errors);
        }

        private static readonly List<string> s_expectedErrors = new()
        {
            "The attribute 't' has invalid value 'd'. The Enumeration constraint failed.",
        };
    }
}

#endif
