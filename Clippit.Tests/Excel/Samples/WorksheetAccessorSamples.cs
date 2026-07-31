using Clippit.Excel;

namespace Clippit.Tests.Excel.Samples
{
    public class WorksheetAccessorSamples() : Clippit.Tests.TestsBase
    {
        private static string GetFilePath(string path) =>
            Path.Combine("../../../Excel/Samples/WorksheetAccessor/", path);

        [Test]
        public void Formulas1()
        {
            var sourceFile = GetFilePath("Formulas1/Formulas.xlsx");
            // Change sheet name in formulas
            using (var streamDoc = new OpenXmlMemoryStreamDocument(OpenXmlPowerToolsDocument.FromFileName(sourceFile)))
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetAccessor.FormulaReplaceSheetName(doc, "Source", "'Source 2'");
                }

                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(TempDir, "FormulasUpdated.xlsx"));
            }

            // Change sheet name in formulas
            using (var streamDoc = new OpenXmlMemoryStreamDocument(OpenXmlPowerToolsDocument.FromFileName(sourceFile)))
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    var sheet = WorksheetAccessor.GetWorksheet(doc, "References");
                    WorksheetAccessor.CopyCellRange(doc, sheet, 1, 1, 7, 5, 4, 8);
                }

                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(TempDir, "FormulasCopied.xlsx"));
            }
        }

        [Test]
        public void PivotTables1()
        {
            // Update an existing pivot table
            var qs = new FileInfo(GetFilePath("PivotTables1/QuarterlySales.xlsx"));
            var qsu = new FileInfo(Path.Combine(TempDir, "QuarterlyPivot.xlsx"));
            var row = 1;
            using (var streamDoc = new OpenXmlMemoryStreamDocument(OpenXmlPowerToolsDocument.FromFileName(qs.FullName)))
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    var sheet = WorksheetAccessor.GetWorksheet(doc, "Range");
                    using (var source = new StreamReader(GetFilePath("PivotTables1/PivotData.txt")))
                    {
                        while (!source.EndOfStream)
                        {
                            var line = source.ReadLine();
                            if (line.Length > 3)
                            {
                                var fields = line.Split(',');
                                var column = 1;
                                foreach (var item in fields)
                                {
                                    if (double.TryParse(item, out var num))
                                        WorksheetAccessor.SetCellValue(doc, sheet, row, column++, num);
                                    else
                                        WorksheetAccessor.SetCellValue(doc, sheet, row, column++, item);
                                }
                            }

                            row++;
                        }
                    }

                    sheet.PutXDocument();
                    WorksheetAccessor.UpdateRangeEndRow(doc, "Sales", row - 1);
                }

                streamDoc.GetModifiedSmlDocument().SaveAs(qsu.FullName);
            }

            // Create from scratch
            row = 1;
            var maxColumn = 1;
            using (var streamDoc = OpenXmlMemoryStreamDocument.CreateSpreadsheetDocument())
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    WorksheetAccessor.CreateDefaultStyles(doc);
                    var sheet = WorksheetAccessor.AddWorksheet(doc, "Range");
                    var ms = new MemorySpreadsheet();
                    var southIndex = WorksheetAccessor.GetStyleIndex(
                        doc,
                        0,
                        8,
                        1,
                        2,
                        new WorksheetAccessor.CellAlignment
                        {
                            HorizontalAlignment = WorksheetAccessor.CellAlignment.Horizontal.Center,
                        },
                        true,
                        false
                    );
                    var gradient = new WorksheetAccessor.GradientFill(90);
                    gradient.AddStop(
                        new WorksheetAccessor.GradientStop(0, new WorksheetAccessor.ColorInfo("FF92D050"))
                    );
                    gradient.AddStop(
                        new WorksheetAccessor.GradientStop(1, new WorksheetAccessor.ColorInfo("FF0070C0"))
                    );
                    var northIndex = WorksheetAccessor.GetStyleIndex(
                        doc,
                        0,
                        WorksheetAccessor.GetFontIndex(
                            doc,
                            new WorksheetAccessor.Font
                            {
                                Italic = true,
                                Size = 8,
                                Color = new WorksheetAccessor.ColorInfo(WorksheetAccessor.ColorInfo.ColorType.Theme, 1),
                                Name = "Times New Roman",
                                Family = 1,
                            }
                        ),
                        WorksheetAccessor.GetFillIndex(doc, gradient),
                        WorksheetAccessor.GetBorderIndex(
                            doc,
                            new WorksheetAccessor.Border
                            {
                                DiagonalDown = true,
                                Diagonal = new WorksheetAccessor.BorderLine(
                                    WorksheetAccessor.BorderLine.LineStyle.Thin,
                                    new WorksheetAccessor.ColorInfo("FF616100")
                                ),
                            }
                        ),
                        null,
                        false,
                        false
                    );
                    WorksheetAccessor.CheckNumberFormat(
                        doc,
                        100,
                        "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)"
                    );
                    var amountIndex = WorksheetAccessor.GetStyleIndex(doc, 100, 0, 0, 0, null, false, false);
                    using (var source = new StreamReader(GetFilePath("PivotTables1/PivotData.txt")))
                    {
                        while (!source.EndOfStream)
                        {
                            var line = source.ReadLine();
                            if (line.Length > 3)
                            {
                                var fields = line.Split(',');
                                var column = 1;
                                foreach (var item in fields)
                                {
                                    if (double.TryParse(item, out var num))
                                    {
                                        if (column == 6)
                                            ms.SetCellValue(row, column++, num, amountIndex);
                                        else
                                            ms.SetCellValue(row, column++, num);
                                    }
                                    else if (item == "Accessories")
                                        ms.SetCellValue(
                                            row,
                                            column++,
                                            item,
                                            WorksheetAccessor.GetStyleIndex(doc, "Good")
                                        );
                                    else if (item == "South")
                                        ms.SetCellValue(row, column++, item, southIndex);
                                    else if (item == "North")
                                        ms.SetCellValue(row, column++, item, northIndex);
                                    else
                                        ms.SetCellValue(row, column++, item);
                                }

                                maxColumn = column - 1;
                            }

                            row++;
                        }
                    }

                    WorksheetAccessor.SetSheetContents(doc, sheet, ms);
                    WorksheetAccessor.SetRange(doc, "Sales", "Range", 1, 1, row - 1, maxColumn);
                    var pivot = WorksheetAccessor.AddWorksheet(doc, "Pivot");
                    WorksheetAccessor.CreatePivotTable(doc, "Sales", pivot);
                    // Configure pivot table rows, columns, data and filters
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Year", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Quarter", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Category", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Product", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddDataValue(doc, pivot, "Amount");
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Region", WorksheetAccessor.PivotAxis.Page);
                }

                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(TempDir, "NewPivot.xlsx"));
            }

            // Add pivot table to existing spreadsheet
            // Demonstrate multiple data fields
            using (
                var streamDoc = new OpenXmlMemoryStreamDocument(
                    OpenXmlPowerToolsDocument.FromFileName(GetFilePath("PivotTables1/QuarterlyUnitSales.xlsx"))
                )
            )
            {
                using (var doc = streamDoc.GetSpreadsheetDocument())
                {
                    var pivot = WorksheetAccessor.AddWorksheet(doc, "Pivot");
                    WorksheetAccessor.CreatePivotTable(doc, "Sales", pivot);
                    // Configure pivot table rows, columns, data and filters
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Year", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Quarter", WorksheetAccessor.PivotAxis.Column);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Category", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Product", WorksheetAccessor.PivotAxis.Row);
                    WorksheetAccessor.AddDataValue(doc, pivot, "Total");
                    WorksheetAccessor.AddDataValue(doc, pivot, "Quantity");
                    WorksheetAccessor.AddDataValue(doc, pivot, "Unit Price");
                    WorksheetAccessor.AddPivotAxis(doc, pivot, "Region", WorksheetAccessor.PivotAxis.Page);
                }

                streamDoc.GetModifiedSmlDocument().SaveAs(Path.Combine(TempDir, "QuarterlyUnitSalesWithPivot.xlsx"));
            }
        }
    }
}
