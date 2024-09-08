﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Excel
{
    public class SmlDataRetriever
    {
        public static XElement RetrieveSheet(SmlDocument smlDoc, string sheetName)
        {
            using var ms = new MemoryStream();
            ms.Write(smlDoc.DocumentByteArray, 0, smlDoc.DocumentByteArray.Length);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            return RetrieveSheet(sDoc, sheetName);
        }

        public static XElement RetrieveSheet(string fileName, string sheetName)
        {
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            return RetrieveSheet(sDoc, sheetName);
        }

        public static XElement RetrieveSheet(SpreadsheetDocument sDoc, string sheetName)
        {
            var wbXdoc = sDoc.WorkbookPart.GetXDocument();
            var sheet = wbXdoc
                .Root.Elements(S.sheets)
                .Elements(S.sheet)
                .FirstOrDefault(s => (string)s.Attribute("name") == sheetName);
            if (sheet == null)
                throw new ArgumentException("Invalid sheet name passed to RetrieveSheet", nameof(sheetName));
            var range = "A1:XFD1048576";
            XlsxTables.ParseRange(range, out var leftColumn, out var topRow, out var rightColumn, out var bottomRow);
            return RetrieveRange(sDoc, sheetName, leftColumn, topRow, rightColumn, bottomRow);
        }

        public static XElement RetrieveRange(SmlDocument smlDoc, string sheetName, string range)
        {
            using var ms = new MemoryStream();
            ms.Write(smlDoc.DocumentByteArray, 0, smlDoc.DocumentByteArray.Length);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            return RetrieveRange(sDoc, sheetName, range);
        }

        public static XElement RetrieveRange(string fileName, string sheetName, string range)
        {
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            return RetrieveRange(sDoc, sheetName, range);
        }

        public static XElement RetrieveRange(SpreadsheetDocument sDoc, string sheetName, string range)
        {
            XlsxTables.ParseRange(range, out var leftColumn, out var topRow, out var rightColumn, out var bottomRow);
            return RetrieveRange(sDoc, sheetName, leftColumn, topRow, rightColumn, bottomRow);
        }

        public static XElement RetrieveRange(
            SmlDocument smlDoc,
            string sheetName,
            int leftColumn,
            int topRow,
            int rightColumn,
            int bottomRow
        )
        {
            using var ms = new MemoryStream();
            ms.Write(smlDoc.DocumentByteArray, 0, smlDoc.DocumentByteArray.Length);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            return RetrieveRange(sDoc, sheetName, leftColumn, topRow, rightColumn, bottomRow);
        }

        public static XElement RetrieveRange(
            string fileName,
            string sheetName,
            int leftColumn,
            int topRow,
            int rightColumn,
            int bottomRow
        )
        {
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            return RetrieveRange(sDoc, sheetName, leftColumn, topRow, rightColumn, bottomRow);
        }

        public static XElement RetrieveRange(
            SpreadsheetDocument sDoc,
            string sheetName,
            int leftColumn,
            int topRow,
            int rightColumn,
            int bottomRow
        )
        {
            var wbXdoc = sDoc.WorkbookPart.GetXDocument();
            var sheet = wbXdoc
                .Root.Elements(S.sheets)
                .Elements(S.sheet)
                .FirstOrDefault(s => (string)s.Attribute("name") == sheetName);
            if (sheet == null)
                throw new ArgumentException("Invalid sheet name passed to RetrieveRange", nameof(sheetName));
            var rId = (string)sheet.Attribute(R.id);
            if (rId == null)
                throw new FileFormatException("Invalid spreadsheet");
            var sheetPart = sDoc.WorkbookPart.GetPartById(rId);
            if (sheetPart == null)
                throw new FileFormatException("Invalid spreadsheet");
            var shXDoc = sheetPart.GetXDocument();

            if (sDoc.WorkbookPart.WorkbookStylesPart == null)
                throw new FileFormatException("Invalid spreadsheet.  No WorkbookStylesPart.");
            var styleXDoc = sDoc.WorkbookPart.WorkbookStylesPart.GetXDocument();

            // if there is no shared string table, sharedStringTable will be null
            // it will only be used if there is a cell type == "s", in which case, referencing this
            // part would indicate an invalid spreadsheet.
            var sharedStringTable = sDoc.WorkbookPart.SharedStringTablePart;

            FixUpCellsThatHaveNoRAtt(shXDoc);

            // assemble the transform
            var sheetData = shXDoc
                .Root.Elements(S.sheetData)
                .Elements(S.row)
                .Select(row =>
                {
                    // filter
                    var ra = (string)row.Attribute("r");
                    if (ra == null)
                        return null;
                    if (!int.TryParse(ra, out var rowNbr))
                        return null;
                    if (rowNbr < topRow)
                        return null;
                    if (rowNbr > bottomRow)
                        return null;

                    var cells = row.Elements(S.c)
                        .Select(cell =>
                        {
                            var cellAddress = (string)cell.Attribute("r");
                            if (cellAddress == null)
                                throw new FileFormatException("Invalid spreadsheet - cell does not have r attribute.");
                            var splitCellAddress = XlsxTables.SplitAddress(cellAddress);
                            var columnAddress = splitCellAddress[0];
                            var columnIndex = XlsxTables.ColumnAddressToIndex(columnAddress);

                            // filter
                            if (columnIndex < leftColumn || columnIndex > rightColumn)
                                return null;

                            var cellType = (string)cell.Attribute("t");
                            string sharedString = null;
                            if (cellType == "s")
                            {
                                var sharedStringBeforeParsing = (string)cell.Element(S.v);
                                if (sharedStringBeforeParsing == null)
                                    sharedStringBeforeParsing = (string)
                                        cell.Elements(S._is).Elements(S.t).FirstOrDefault();
                                if (sharedStringBeforeParsing == null)
                                    throw new FileFormatException("Invalid document");
                                if (!int.TryParse(sharedStringBeforeParsing, out var sharedStringIndex))
                                    throw new FileFormatException("Invalid document");
                                XElement sharedStringElement = null;
                                if (sharedStringTable == null)
                                    throw new FileFormatException(
                                        "Invalid spreadsheet.  Shared string, but no Shared String Part."
                                    );
                                sharedStringElement = sharedStringTable
                                    .GetXDocument()
                                    .Root.Elements(S.si)
                                    .Skip(sharedStringIndex)
                                    .FirstOrDefault();
                                if (sharedStringElement == null)
                                    throw new FileFormatException(
                                        "Invalid spreadsheet.  Shared string reference not valid."
                                    );
                                sharedString = sharedStringElement.Descendants(S.t).StringConcatenate(e => (string)e);
                            }

                            if (sharedString != null)
                            {
                                var cellProps = GetCellProps_NotInTable(sDoc, styleXDoc, cell);
                                var value = sharedString;
                                string displayValue;
                                string color = null;
                                if (cellProps != null)
                                    displayValue = SmlCellFormatter.FormatCell(
                                        (string)cellProps.Attribute("formatCode"),
                                        value,
                                        out color
                                    );
                                else
                                    displayValue = value;
                                var newCell1 = new XElement(
                                    "Cell",
                                    new XAttribute("Ref", (string)cell.Attribute("r")),
                                    new XAttribute("ColumnId", columnAddress),
                                    new XAttribute("ColumnNumber", columnIndex),
                                    cell.Attribute("f") != null
                                        ? new XAttribute("Formula", (string)cell.Attribute("f"))
                                        : null,
                                    cell.Attribute("s") != null
                                        ? new XAttribute("Style", (string)cell.Attribute("s"))
                                        : null,
                                    cell.Attribute("t") != null
                                        ? new XAttribute("Type", (string)cell.Attribute("t"))
                                        : null,
                                    cellProps,
                                    new XElement("Value", value),
                                    new XElement("DisplayValue", displayValue),
                                    color != null ? new XElement("DisplayColor", color) : null
                                );
                                return newCell1;
                            }
                            else
                            {
                                var type = (string)cell.Attribute("t");
                                var value = new XElement("Value", cell.Value);
                                if (type != null && type == "inlineStr")
                                {
                                    type = "s";
                                }
                                XAttribute typeAttr = null;
                                if (type != null)
                                    typeAttr = new XAttribute("Type", type);

                                var cellProps = GetCellProps_NotInTable(sDoc, styleXDoc, cell);
                                string displayValue;
                                string color = null;
                                if (cellProps != null)
                                    displayValue = SmlCellFormatter.FormatCell(
                                        (string)cellProps.Attribute("formatCode"),
                                        cell.Value,
                                        out color
                                    );
                                else
                                    displayValue = displayValue = SmlCellFormatter.FormatCell(
                                        "General",
                                        cell.Value,
                                        out color
                                    );
                                var newCell2 = new XElement(
                                    "Cell",
                                    new XAttribute("Ref", (string)cell.Attribute("r")),
                                    new XAttribute("ColumnId", columnAddress),
                                    new XAttribute("ColumnNumber", columnIndex),
                                    typeAttr,
                                    cell.Attribute("f") != null
                                        ? new XAttribute("Formula", (string)cell.Attribute("f"))
                                        : null,
                                    cell.Attribute("s") != null
                                        ? new XAttribute("Style", (string)cell.Attribute("s"))
                                        : null,
                                    cellProps,
                                    value,
                                    new XElement("DisplayValue", displayValue),
                                    color != null ? new XElement("DisplayColor", color) : null
                                );
                                return newCell2;
                            }
                        });
                    var dataRow = new XElement(
                        "Row",
                        row.Attribute("r") != null ? new XAttribute("RowNumber", (int)row.Attribute("r")) : null,
                        cells
                    );
                    return dataRow;
                });

            var dataProps = GetDataProps(shXDoc);
            var data = new XElement("Data", dataProps, sheetData);
            return data;
        }

        // Sometimes encounter cells that have no r attribute, so infer it if possible.
        // These are invalid spreadsheets, but attempt to get the data anyway.
        private static void FixUpCellsThatHaveNoRAtt(XDocument shXDoc)
        {
            // if there are any rows that have all cells with no r attribute, then fix them up
            var invalidRows = shXDoc
                .Descendants(S.row)
                .Where(r => !r.Elements(S.c).Any(c => c.Attribute("r") != null))
                .ToList();

            foreach (var row in invalidRows)
            {
                var rowNumberStr = (string)row.Attribute("r");
                var colNumber = 0;
                foreach (var cell in row.Elements(S.c))
                {
                    var newCellRef = XlsxTables.IndexToColumnAddress(colNumber) + rowNumberStr;
                    cell.Add(new XAttribute("r", newCellRef));
                }
            }

            // repeat iteratively until no further fixes can be made
            while (true)
            {
                var invalidCells = shXDoc.Descendants(S.c).Where(c => c.Attribute("r") == null).ToList();

                var didFixup = false;
                foreach (var cell in invalidCells)
                {
                    var followingCell = cell.ElementsAfterSelf(S.c).FirstOrDefault();
                    if (followingCell != null)
                    {
                        var followingR = (string)followingCell.Attribute("r");
                        if (followingR != null)
                        {
                            var spl = XlsxTables.SplitAddress(followingR);
                            var colIdxFollowing = XlsxTables.ColumnAddressToIndex(spl[0]);
                            var newRef = XlsxTables.IndexToColumnAddress(colIdxFollowing - 1) + spl[1];
                            cell.Add(new XAttribute("r", newRef));
                            didFixup = true;
                        }
                        else
                        {
                            didFixup = FixUpBasedOnPrecedingCell(didFixup, cell);
                        }
                    }
                    else
                    {
                        didFixup = FixUpBasedOnPrecedingCell(didFixup, cell);
                    }
                }
                if (!didFixup)
                    break;
            }
        }

        private static bool FixUpBasedOnPrecedingCell(bool didFixup, XElement cell)
        {
            var precedingCell = GetPrevousElement(cell);
            if (precedingCell != null)
            {
                var precedingR = (string)precedingCell.Attribute("r");
                if (precedingR != null)
                {
                    var spl = XlsxTables.SplitAddress(precedingR);
                    var colIdxFollowing = XlsxTables.ColumnAddressToIndex(spl[0]);
                    var newRef = XlsxTables.IndexToColumnAddress(colIdxFollowing + 1) + spl[1];
                    cell.Add(new XAttribute("r", newRef));
                    didFixup = true;
                }
            }
            return didFixup;
        }

        private static XElement GetPrevousElement(XElement element)
        {
            XElement previousElement = null;
            XNode currentNode = element;
            while (true)
            {
                if (currentNode.PreviousNode == null)
                    return null;
                previousElement = currentNode.PreviousNode as XElement;
                if (previousElement != null)
                    return previousElement;
                currentNode = currentNode.PreviousNode;
            }
        }

        public static XElement RetrieveTable(SmlDocument smlDoc, string sheetName, string tableName)
        {
            using var ms = new MemoryStream();
            ms.Write(smlDoc.DocumentByteArray, 0, smlDoc.DocumentByteArray.Length);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            return RetrieveTable(sDoc, tableName);
        }

        public static XElement RetrieveTable(string fileName, string tableName)
        {
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            return RetrieveTable(sDoc, tableName);
        }

        public static XElement RetrieveTable(SpreadsheetDocument sDoc, string tableName)
        {
            var table = sDoc.Table(tableName);
            if (table == null)
                throw new ArgumentException("Table not found", nameof(tableName));

            var styleXDoc = sDoc.WorkbookPart.WorkbookStylesPart.GetXDocument();
            var r = table.Ref;
            XlsxTables.ParseRange(r, out var leftColumn, out var topRow, out var rightColumn, out var bottomRow);
            var shXDoc = table.Parent.GetXDocument();

            FixUpCellsThatHaveNoRAtt(shXDoc);

            // assemble the transform
            var columns = new XElement(
                "Columns",
                table
                    .TableColumns()
                    .Select(tc =>
                    {
                        var colXElement = new XElement(
                            "Column",
                            tc.Name != null ? new XAttribute("Name", tc.Name) : null,
                            tc.UniqueName != null ? new XAttribute("UniqueName", tc.UniqueName) : null,
                            new XAttribute("ColumnIndex", tc.ColumnIndex),
                            new XAttribute("Id", tc.Id),
                            tc.DataDxfId != null ? new XAttribute("DataDxfId", tc.DataDxfId) : null,
                            tc.QueryTableFieldId != null
                                ? new XAttribute("QueryTableFieldId", tc.QueryTableFieldId)
                                : null
                        );
                        return colXElement;
                    })
            );

            var dataProps = GetDataProps(shXDoc);
            var data = new XElement(
                "Data",
                dataProps,
                table
                    .TableRows()
                    .Select(tr =>
                    {
                        if (!int.TryParse(tr.Row.RowId, out var rowRef))
                            throw new FileFormatException("Invalid spreadsheet");

                        // filter
                        if (rowRef < topRow || rowRef > bottomRow)
                            return null;

                        var cellData = tr
                            .Row.Cells()
                            .Select(tc =>
                            {
                                // filter
                                var columnIndex = tc.ColumnIndex;
                                if (columnIndex < leftColumn || columnIndex > rightColumn)
                                    return null;

                                var cellProps = GetCellProps_InTable(sDoc, styleXDoc, table, tc);
                                if (tc.SharedString != null)
                                {
                                    string displayValue;
                                    string color = null;
                                    if (cellProps != null)
                                        displayValue = SmlCellFormatter.FormatCell(
                                            (string)cellProps.Attribute("formatCode"),
                                            tc.SharedString,
                                            out color
                                        );
                                    else
                                        displayValue = tc.SharedString;
                                    var newCell1 = new XElement(
                                        "Cell",
                                        tc.CellElement != null
                                            ? new XAttribute("Ref", (string)tc.CellElement.Attribute("r"))
                                            : null,
                                        tc.ColumnAddress != null ? new XAttribute("ColumnId", tc.ColumnAddress) : null,
                                        new XAttribute("ColumnNumber", tc.ColumnIndex),
                                        tc.Type != null ? new XAttribute("Type", "s") : null,
                                        tc.Formula != null ? new XAttribute("Formula", tc.Formula) : null,
                                        tc.Style != null ? new XAttribute("Style", tc.Style) : null,
                                        cellProps,
                                        new XElement("Value", tc.SharedString),
                                        new XElement("DisplayValue", displayValue),
                                        color != null ? new XElement("DisplayColor", color) : null
                                    );
                                    return newCell1;
                                }
                                else
                                {
                                    XAttribute type = null;
                                    if (tc.Type != null)
                                    {
                                        if (tc.Type == "inlineStr")
                                            type = new XAttribute("Type", "s");
                                        else
                                            type = new XAttribute("Type", tc.Type);
                                    }
                                    string displayValue;
                                    string color = null;
                                    if (cellProps != null)
                                        displayValue = SmlCellFormatter.FormatCell(
                                            (string)cellProps.Attribute("formatCode"),
                                            tc.Value,
                                            out color
                                        );
                                    else
                                        displayValue = SmlCellFormatter.FormatCell("General", tc.Value, out color);
                                    var newCell = new XElement(
                                        "Cell",
                                        tc.CellElement != null
                                            ? new XAttribute("Ref", (string)tc.CellElement.Attribute("r"))
                                            : null,
                                        tc.ColumnAddress != null ? new XAttribute("ColumnId", tc.ColumnAddress) : null,
                                        new XAttribute("ColumnNumber", tc.ColumnIndex),
                                        type,
                                        tc.Formula != null ? new XAttribute("Formula", tc.Formula) : null,
                                        tc.Style != null ? new XAttribute("Style", tc.Style) : null,
                                        cellProps,
                                        new XElement("Value", tc.Value),
                                        new XElement("DisplayValue", displayValue),
                                        color != null ? new XElement("DisplayColor", color) : null
                                    );
                                    return newCell;
                                }
                            });
                        var rowProps = GetRowProps(tr.Row.RowElement);
                        var newRow = new XElement("Row", rowProps, new XAttribute("RowNumber", tr.Row.RowId), cellData);
                        return newRow;
                    })
            );

            var tableProps = GetTableProps(table);
            var tableXml = new XElement(
                "Table",
                tableProps,
                table.TableName != null ? new XAttribute("TableName", table.TableName) : null,
                table.DisplayName != null ? new XAttribute("DisplayName", table.DisplayName) : null,
                table.Ref != null ? new XAttribute("Ref", table.Ref) : null,
                table.HeaderRowCount != null ? new XAttribute("HeaderRowCount", table.HeaderRowCount) : null,
                table.TotalsRowCount != null ? new XAttribute("TotalsRowCount", table.TotalsRowCount) : null,
                columns,
                data
            );
            return tableXml;
        }

        public static string[] SheetNames(SmlDocument smlDoc)
        {
            using var ms = new MemoryStream();
            ms.Write(smlDoc.DocumentByteArray, 0, smlDoc.DocumentByteArray.Length);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            return SheetNames(sDoc);
        }

        public static string[] SheetNames(string fileName)
        {
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            return SheetNames(sDoc);
        }

        public static string[] SheetNames(SpreadsheetDocument sDoc)
        {
            var workbookXDoc = sDoc.WorkbookPart.GetXDocument();
            var sheetNames = workbookXDoc
                .Root.Elements(S.sheets)
                .Elements(S.sheet)
                .Attributes("name")
                .Select(a => (string)a)
                .ToArray();
            return sheetNames;
        }

        public static string[] TableNames(SmlDocument smlDoc)
        {
            using var ms = new MemoryStream();
            ms.Write(smlDoc.DocumentByteArray, 0, smlDoc.DocumentByteArray.Length);
            using var sDoc = SpreadsheetDocument.Open(ms, false);
            return TableNames(sDoc);
        }

        public static string[] TableNames(string fileName)
        {
            using var sDoc = SpreadsheetDocument.Open(fileName, false);
            return TableNames(sDoc);
        }

        public static string[] TableNames(SpreadsheetDocument sDoc)
        {
            var workbookXDoc = sDoc.WorkbookPart.GetXDocument();
            var sheets = workbookXDoc.Root.Elements(S.sheets).Elements(S.sheet);
            var tableNames = sheets
                .Select(sh =>
                {
                    var rId = (string)sh.Attribute(R.id);
                    var sheetPart = sDoc.WorkbookPart.GetPartById(rId);
                    var sheetXDoc = sheetPart.GetXDocument();
                    var tableParts = sheetXDoc.Root.Element(S.tableParts);
                    if (tableParts == null)
                        return new List<string>();
                    var tableNames2 = tableParts
                        .Elements(S.tablePart)
                        .Select(tp =>
                        {
                            var tpRId = (string)tp.Attribute(R.id);
                            var tpart = sheetPart.GetPartById(tpRId);
                            var tpxd = tpart.GetXDocument();
                            var name = (string)tpxd.Root.Attribute("name");
                            return name;
                        })
                        .ToList();
                    return tableNames2;
                })
                .SelectMany(m => m)
                .ToArray();
            return tableNames;
        }

        private static XElement GetTableProps(Table table)
        {
            var tableProps = new XElement("TableProps");
            var tableStyleInfo = table.TableStyleInfo;
            if (tableStyleInfo != null)
            {
                var newTableStyleInfo = TransformRemoveNamespace(tableStyleInfo);
                tableProps.Add(newTableStyleInfo);
            }

            if (!tableProps.HasElements && !tableProps.HasElements)
                tableProps = null;
            return tableProps;
        }

        private static XElement GetDataProps(XDocument shXDoc)
        {
            var sheetFormatPr = shXDoc.Root.Element(S.sheetFormatPr);
            if (sheetFormatPr != null && sheetFormatPr.Attribute("defaultColWidth") == null)
                sheetFormatPr.Add(new XAttribute("defaultColWidth", "9.25"));
            if (sheetFormatPr != null && sheetFormatPr.Attribute("defaultRowHeight") == null)
                sheetFormatPr.Add(new XAttribute("defaultRowHeight", "14.25"));

            var mergeCells = TransformRemoveNamespace(shXDoc.Root.Element(S.mergeCells));

            var dataProps = new XElement(
                "DataProps",
                TransformRemoveNamespace(sheetFormatPr),
                TransformRemoveNamespace(shXDoc.Root.Element(S.cols)),
                mergeCells
            );

            if (!dataProps.HasAttributes && !dataProps.HasElements)
                dataProps = null;
            return dataProps;
        }

        private static XElement GetRowProps(XElement rowElement)
        {
            var rowProps = new XElement("RowProps");
            var ht = rowElement.Attribute("ht");
            if (ht != null)
                rowProps.Add(ht);
            var dyDescent = rowElement.Attribute(x14ac + "dyDescent");
            if (dyDescent != null)
                rowProps.Add(new XAttribute("dyDescent", (string)dyDescent));

            if (!rowProps.HasAttributes && !rowProps.HasElements)
                rowProps = null;
            return rowProps;
        }

        private static XElement GetCellProps_NotInTable(SpreadsheetDocument sDoc, XDocument styleXDoc, XElement cell)
        {
            var cellProps = new XElement("CellProps");
            var style = (int?)cell.Attribute("s");
            if (style == null)
                return cellProps;

            var xf = styleXDoc.Root.Elements(S.cellXfs).Elements(S.xf).Skip((int)style).FirstOrDefault();

            var numFmtId = (int?)xf.Attribute("numFmtId");
            if (numFmtId != null)
                AddNumFmtIdAndFormatCode(styleXDoc, cellProps, numFmtId);

            var masterXfId = (int?)xf.Attribute("xfId");
            if (masterXfId != null)
            {
                var masterXf = styleXDoc
                    .Root.Elements(S.cellStyleXfs)
                    .Elements(S.xf)
                    .Skip((int)masterXfId)
                    .FirstOrDefault();
                if (masterXf != null)
                    AddFormattingToCellProps(styleXDoc, cellProps, masterXf);
            }

            AddFormattingToCellProps(styleXDoc, cellProps, xf);
            AugmentAndCleanUpProps(cellProps);

            if (!cellProps.HasElements && !cellProps.HasAttributes)
                return null;
            return cellProps;
        }

        private static XElement GetCellProps_InTable(
            SpreadsheetDocument sDoc,
            XDocument styleXDoc,
            Table table,
            Cell tc
        )
        {
            var style = tc.Style;
            if (style == null)
                return null;

            var colIdStr = tc.ColumnAddress;
            var colNbr = XlsxTables.ColumnAddressToIndex(colIdStr);
            var column = table.TableColumns().FirstOrDefault(z => z.ColumnNumber == colNbr);
            if (column == null)
                throw new FileFormatException("Invalid spreadsheet");

            var cellProps = new XElement("CellProps");
            var d = column.DataDxfId;
            if (d != null)
            {
                var dataDxf = styleXDoc.Root.Elements(S.dxfs).Elements(S.dxf).Skip((int)d).FirstOrDefault();
                if (dataDxf == null)
                    throw new FileFormatException("Invalid spreadsheet");

                var numFmt = dataDxf.Element(S.numFmt);
                if (numFmt != null)
                {
                    var numFmtId = (int?)numFmt.Attribute("numFmtId");
                    if (numFmtId != null)
                        cellProps.Add(new XAttribute("numFmtId", numFmtId));
                    var formatCode = (string)numFmt.Attribute("formatCode");
                    if (formatCode != null)
                        cellProps.Add(new XAttribute("formatCode", formatCode));
                }
            }

            var xf = styleXDoc.Root.Elements(S.cellXfs).Elements(S.xf).Skip((int)style).FirstOrDefault();
            if (xf == null)
                throw new FileFormatException("Invalid spreadsheet");

            // if xf has different numFmtId, then replace the ones from the table definition
            var numFmtId2 = (int?)xf.Attribute("numFmtId");
            if (numFmtId2 != null)
                AddNumFmtIdAndFormatCode(styleXDoc, cellProps, numFmtId2);

            var masterXfId = (int?)xf.Attribute("xfId");
            if (masterXfId != null)
            {
                var masterXf = styleXDoc
                    .Root.Elements(S.cellStyleXfs)
                    .Elements(S.xf)
                    .Skip((int)masterXfId)
                    .FirstOrDefault();
                if (masterXf != null)
                    AddFormattingToCellProps(styleXDoc, cellProps, masterXf);
            }

            AddFormattingToCellProps(styleXDoc, cellProps, xf);
            AugmentAndCleanUpProps(cellProps);

            if (!cellProps.HasElements && !cellProps.HasAttributes)
                return null;
            return cellProps;
        }

        private static void AddNumFmtIdAndFormatCode(XDocument styleXDoc, XElement props, int? numFmtId)
        {
            var existingNumFmtId = props.Attribute("numFmtId");
            if (existingNumFmtId != null)
                existingNumFmtId.Value = numFmtId.ToString();
            else
                props.Add(new XAttribute("numFmtId", numFmtId));

            var numFmt = styleXDoc
                .Root.Elements(S.numFmts)
                .Elements(S.numFmt)
                .FirstOrDefault(z => (int)z.Attribute("numFmtId") == numFmtId);

            if (numFmt == null)
            {
                var formatCode = GetFormatCodeFromFmtId((int)numFmtId);
                if (formatCode != null)
                {
                    var existingFormatCode = props.Attribute("formatCode");
                    if (existingFormatCode != null)
                        existingFormatCode.Value = formatCode;
                    else
                        props.Add(new XAttribute("formatCode", formatCode));
                }
            }
            else
            {
                var formatCode = (string)numFmt.Attribute("formatCode");
                if (formatCode != null)
                {
                    var existingFormatCode = props.Attribute("formatCode");
                    if (existingFormatCode != null)
                        existingFormatCode.Value = formatCode;
                    else
                        props.Add(new XAttribute("formatCode", formatCode));
                }
            }
        }

        private static void AddFormattingToCellProps(XDocument styleXDoc, XElement props, XElement xf)
        {
            MoveBooleanAttribute(props, xf, "applyAlignment");
            MoveBooleanAttribute(props, xf, "applyBorder");
            MoveBooleanAttribute(props, xf, "applyFill");
            MoveBooleanAttribute(props, xf, "applyFont");
            MoveBooleanAttribute(props, xf, "applyNumberFormat");

            var borderId = (int?)xf.Attribute("borderId");
            var fillId = (int?)xf.Attribute("fillId");
            var fontId = (int?)xf.Attribute("fontId");

            if (fontId != null)
            {
                var fontElement = styleXDoc.Root.Elements(S.fonts).Elements(S.font).Skip((int)fontId).FirstOrDefault();
                if (fontElement != null)
                {
                    var newFontElement = (XElement)TransformRemoveNamespace(fontElement);
                    AddOrReplaceElement(props, newFontElement);
                }
            }

            if (fillId != null)
            {
                var fillElement = styleXDoc.Root.Elements(S.fills).Elements(S.fill).Skip((int)fillId).FirstOrDefault();
                if (fillElement != null)
                {
                    var newFillElement = (XElement)TransformRemoveNamespace(fillElement);
                    AddOrReplaceElement(props, newFillElement);
                }
            }

            if (borderId != null)
            {
                var borderElement = styleXDoc
                    .Root.Elements(S.borders)
                    .Elements(S.border)
                    .Skip((int)borderId)
                    .FirstOrDefault();
                if (borderElement != null)
                {
                    var newborderElement = (XElement)TransformRemoveNamespace(borderElement);
                    AddOrReplaceElement(props, newborderElement);
                }
            }

            if (xf.Element(S.alignment) != null)
            {
                var newAlignmentElement = (XElement)TransformRemoveNamespace(xf.Element(S.alignment));
                AddOrReplaceElement(props, newAlignmentElement);
            }
        }

        private static void MoveBooleanAttribute(XElement props, XElement xf, XName attributeName)
        {
            var attrValue = ConvertAttributeToBool(xf.Attribute(attributeName));
            if (attrValue)
            {
                if (props.Attribute(attributeName) == null)
                    props.Add(new XAttribute(attributeName, attrValue ? "1" : "0"));
                else
                    props.Attribute(attributeName).Value = attrValue ? "1" : "0";
            }
        }

        public static string[] IndexedColors = new string[]
        {
            "00000000",
            "00FFFFFF",
            "00FF0000",
            "0000FF00",
            "000000FF",
            "00FFFF00",
            "00FF00FF",
            "0000FFFF",
            "00000000",
            "00FFFFFF",
            "00FF0000",
            "0000FF00",
            "000000FF",
            "00FFFF00",
            "00FF00FF",
            "0000FFFF",
            "00800000",
            "00008000",
            "00000080",
            "00808000",
            "00800080",
            "00008080",
            "00C0C0C0",
            "00808080",
            "009999FF",
            "00993366",
            "00FFFFCC",
            "00CCFFFF",
            "00660066",
            "00FF8080",
            "000066CC",
            "00CCCCFF",
            "00000080",
            "00FF00FF",
            "00FFFF00",
            "0000FFFF",
            "00800080",
            "00800000",
            "00008080",
            "000000FF",
            "0000CCFF",
            "00CCFFFF",
            "00CCFFCC",
            "00FFFF99",
            "0099CCFF",
            "00FF99CC",
            "00CC99FF",
            "00FFCC99",
            "003366FF",
            "0033CCCC",
            "0099CC00",
            "00FFCC00",
            "00FF9900",
            "00FF6600",
            "00666699",
            "00969696",
            "00003366",
            "00339966",
            "00003300",
            "00333300",
            "00993300",
            "00993366",
            "00333399",
            "00333333",
            "System Foreground",
            "System Background",
        };

        private static readonly string[] FontFamilyList = new string[]
        {
            "Not applicable",
            "Roman",
            "Swiss",
            "Modern",
            "Script",
            "Decorative",
        };

        private static void AugmentAndCleanUpProps(XElement props)
        {
            foreach (var color in props.Descendants("color").Where(c => c.Attribute("indexed") != null).ToList())
            {
                var idx = (int)color.Attribute("indexed");
                if (idx < IndexedColors.Length)
                {
                    color.Add(new XAttribute("val", IndexedColors[idx]));
                }
                color.Attribute("indexed").Remove();
            }
            foreach (var family in props.Descendants("family").ToList())
            {
                var fam = (int?)family.Attribute("val");
                if (fam != null)
                {
                    if (fam < FontFamilyList.Length)
                    {
                        family.Attribute("val").Remove();
                        family.Add(new XAttribute("val", FontFamilyList[(int)fam]));
                    }
                }
            }
            foreach (var border in props.Descendants("border").ToList())
            {
                RemoveIfEmpty(border.Element("left"));
                RemoveIfEmpty(border.Element("right"));
                RemoveIfEmpty(border.Element("top"));
                RemoveIfEmpty(border.Element("bottom"));
                RemoveIfEmpty(border.Element("diagonal"));
                if (!border.HasAttributes && !border.HasElements)
                    border.Remove();
            }
            foreach (var fill in props.Descendants("fill").ToList())
            {
                fill.Elements("patternFill").Where(pf => (string)pf.Attribute("patternType") == "none").Remove();
                if (!fill.HasAttributes && !fill.HasElements)
                    fill.Remove();
            }
        }

        private static void RemoveIfEmpty(XElement xElement)
        {
            if (xElement == null)
                return;
            if (!xElement.HasAttributes && !xElement.HasElements)
                xElement.Remove();
        }

        private static object TransformRemoveNamespace(XNode node)
        {
            return node switch
            {
                null => null,
                XElement element => new XElement(
                    element.Name.LocalName,
                    element
                        .Attributes()
                        .Select(a => new XAttribute(a.Name.LocalName, (string)a))
                        .OrderBy(a => a.Name.LocalName),
                    element.Nodes().Select(n => TransformRemoveNamespace(n))
                ),
                _ => node,
            };
        }

        private static string GetFormatCodeFromFmtId(int fmtId)
        {
            return fmtId switch
            {
                0 => "General",
                1 => "0",
                2 => "0.00",
                3 => "#,##0",
                4 => "#,##0.00",
                9 => "0%",
                10 => "0.00%",
                11 => "0.00E+00",
                12 => "# ?/?",
                13 => "# ??/??",
                14 => "mm-dd-yy",
                15 => "d-mmm-yy",
                16 => "d-mmm",
                17 => "mmm-yy",
                18 => "h:mm AM/PM",
                19 => "h:mm:ss AM/PM",
                20 => "h:mm",
                21 => "h:mm:ss",
                22 => "22 m/d/yy h:mm",
                37 => "#,##0 ;(#,##0)",
                38 => "#,##0 ;[Red](#,##0)",
                39 => "#,##0.00;(#,##0.00)",
                40 => "#,##0.00;[Red](#,##0.00)",
                45 => "mm:ss",
                46 => "[h]:mm:ss",
                47 => "mmss.0",
                48 => "##0.0E+0",
                49 => "@",
                _ => null,
            };
        }

        private static void AddOrReplaceElement(XElement props, XName childElementName, int value)
        {
            var existingElement = props.Element(childElementName);
            if (existingElement != null)
                existingElement.ReplaceWith(new XElement(childElementName, new XAttribute("Val", value)));
            else
                props.Add(new XElement(childElementName, new XAttribute("Val", value)));
        }

        private static void AddOrReplaceElement(XElement props, XName childElementName, string value)
        {
            var existingElement = props.Element(childElementName);
            if (existingElement != null)
                existingElement.ReplaceWith(new XElement(childElementName, new XAttribute("Val", value)));
            else
                props.Add(new XElement(childElementName, new XAttribute("Val", value)));
        }

        private static void AddOrReplaceElement(XElement props, XElement element)
        {
            var existingElement = props.Element(element.Name);
            if (existingElement != null)
                existingElement.ReplaceWith(element);
            else
                props.Add(element);
        }

        private static bool ConvertAttributeToBool(XAttribute xAttribute)
        {
            var applyNumberFormatStr = (string)xAttribute;
            var returnValue = false;
            if (applyNumberFormatStr != null)
            {
                if (applyNumberFormatStr == "1")
                    returnValue = true;
                if (applyNumberFormatStr.Substring(0, 1).ToUpper() == "T")
                    returnValue = true;
            }
            return returnValue;
        }

        private static readonly XNamespace x14ac = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
    }
}
