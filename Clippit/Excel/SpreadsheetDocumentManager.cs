// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Excel
{
    /// <summary>
    /// Manages SpreadsheetDocument content
    /// </summary>
    public class SpreadsheetDocumentManager
    {
        private static readonly XNamespace ns;
        private static readonly XNamespace relationshipsns;
        private static int headerRow = 1;

        static SpreadsheetDocumentManager()
        {
            ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            relationshipsns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        }

        /// <summary>
        /// Creates a spreadsheet document from a value table
        /// </summary>
        /// <param name="filePath">Path to store the document</param>
        /// <param name="headerList">Contents of first row (header)</param>
        /// <param name="valueTable">Contents of data</param>
        /// <param name="initialRow">Row to start copying data from</param>
        /// <returns></returns>
        public static void Create(
            SpreadsheetDocument document,
            List<string> headerList,
            string[][] valueTable,
            int initialRow
        )
        {
            headerRow = initialRow;

            //Creates a worksheet with given data
            var worksheet = WorksheetAccessor.Create(document, headerList, valueTable, headerRow);
        }

        /// <summary>
        /// Creates a spreadsheet document with a chart from a value table
        /// </summary>
        /// <param name="filePath">Path to store the document</param>
        /// <param name="headerList">Contents of first row (header)</param>
        /// <param name="valueTable">Contents of data</param>
        /// <param name="chartType">Chart type</param>
        /// <param name="categoryColumn">Column to use as category for charting</param>
        /// <param name="columnsToChart">Columns to use as data series</param>
        /// <param name="initialRow">Row index to start copying data</param>
        /// <returns>SpreadsheetDocument</returns>
        //public static void Create(SpreadsheetDocument document, List<string> headerList, string[][] valueTable, ChartType chartType, string categoryColumn, List<string> columnsToChart, int initialRow)
        //{
        //    headerRow = initialRow;

        //    //Creates worksheet with data
        //    WorksheetPart worksheet = WorksheetAccessor.Create(document, headerList, valueTable, headerRow);
        //    //Creates chartsheet with given series and category
        //    string sheetName = GetSheetName(worksheet, document);
        //    ChartsheetPart chartsheet =
        //        ChartsheetAccessor.Create(document,
        //            chartType,
        //            GetValueReferences(sheetName, categoryColumn, headerList, columnsToChart, valueTable),
        //            GetHeaderReferences(sheetName, categoryColumn, headerList, columnsToChart, valueTable),
        //            GetCategoryReference(sheetName, categoryColumn, headerList, valueTable)
        //        );
        //}

        /// <summary>
        /// Gets the internal name of a worksheet from a document
        /// </summary>
        private static string GetSheetName(WorksheetPart worksheet, SpreadsheetDocument document)
        {
            //Gets the id of worksheet part
            var partId = document.WorkbookPart.GetIdOfPart(worksheet);
            var workbookDocument = document.WorkbookPart.GetXDocument();
            //Gets the name from sheet tag related to worksheet
            var sheetName = workbookDocument
                .Root.Element(ns + "sheets")
                .Elements(ns + "sheet")
                .Where(t => t.Attribute(relationshipsns + "id").Value == partId)
                .First()
                .Attribute("name")
                .Value;
            return sheetName;
        }

        /// <summary>
        /// Gets the range reference for category
        /// </summary>
        /// <param name="sheetName">worksheet to take data from</param>
        /// <param name="headerColumn">name of column used as category</param>
        /// <param name="headerList">column names from data</param>
        /// <param name="valueTable">Data values</param>
        /// <returns></returns>
        private static string GetCategoryReference(
            string sheetName,
            string headerColumn,
            List<string> headerList,
            string[][] valueTable
        )
        {
            var categoryColumn = headerList.IndexOf(headerColumn.ToUpper()) + 1;
            var numRows = valueTable.GetLength(0);

            return GetRangeReference(sheetName, categoryColumn, headerRow + 1, categoryColumn, numRows + headerRow);
        }

        /// <summary>
        /// Gets a list of range references for each of the series headers
        /// </summary>
        /// <param name="sheetName">worksheet to take data from</param>
        /// <param name="headerColumn">name of column used as category</param>
        /// <param name="headerList">column names from data</param>
        /// <param name="valueTable">Data values</param>
        /// <param name="colsToChart">Columns used as data series</param>
        /// <returns></returns>
        private static List<string> GetHeaderReferences(
            string sheetName,
            string headerColumn,
            List<string> headerList,
            List<string> colsToChart,
            string[][] valueTable
        )
        {
            var valueReferenceList = new List<string>();

            foreach (var column in colsToChart)
            {
                valueReferenceList.Add(
                    GetRangeReference(sheetName, headerList.IndexOf(column.ToUpper()) + 1, headerRow)
                );
            }
            return valueReferenceList;
        }

        /// <summary>
        /// Gets a list of range references for each of the series values
        /// </summary>
        /// <param name="sheetName">worksheet to take data from</param>
        /// <param name="headerColumn">name of column used as category</param>
        /// <param name="headerList">column names from data</param>
        /// <param name="valueTable">Data values</param>
        /// <param name="colsToChart">Columns used as data series</param>
        /// <returns></returns>
        private static List<string> GetValueReferences(
            string sheetName,
            string headerColumn,
            List<string> headerList,
            List<string> colsToChart,
            string[][] valueTable
        )
        {
            var valueReferenceList = new List<string>();
            var numRows = valueTable.GetLength(0);

            foreach (var column in colsToChart)
            {
                var dataColumn = headerList.IndexOf(column.ToUpper()) + 1;
                valueReferenceList.Add(
                    GetRangeReference(sheetName, dataColumn, headerRow + 1, dataColumn, numRows + headerRow)
                );
            }
            return valueReferenceList;
        }

        /// <summary>
        /// Gets a formatted representation of a cell range from a worksheet
        /// </summary>
        private static string GetRangeReference(string worksheet, int column, int row)
        {
            return $"{worksheet}!{WorksheetAccessor.GetColumnId(column)}{row}";
        }

        /// <summary>
        /// Gets a formatted representation of a cell range from a worksheet
        /// </summary>
        private static string GetRangeReference(
            string worksheet,
            int startColumn,
            int startRow,
            int endColumn,
            int endRow
        )
        {
            return $"{worksheet}!{WorksheetAccessor.GetColumnId(startColumn)}{startRow}:{WorksheetAccessor.GetColumnId(endColumn)}{endRow}";
        }

        /// <summary>
        /// Creates an empty (base) workbook document
        /// </summary>
        /// <returns></returns>
        private static XDocument CreateEmptyWorkbook()
        {
            var document = new XDocument(
                new XElement(
                    ns + "workbook",
                    new XAttribute("xmlns", ns),
                    new XAttribute(XNamespace.Xmlns + "r", relationshipsns),
                    new XElement(ns + "sheets")
                )
            );

            return document;
        }
    }
}
