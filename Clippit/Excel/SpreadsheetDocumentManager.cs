// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;

namespace Clippit.Excel
{
    /// <summary>
    /// Manages SpreadsheetDocument content
    /// </summary>
    public class SpreadsheetDocumentManager
    {
        /// <summary>
        /// Creates a spreadsheet document from a value table
        /// </summary>
        /// <param name="document">The SpreadsheetDocument to write to</param>
        /// <param name="headerList">Contents of first row (header)</param>
        /// <param name="valueTable">Contents of data</param>
        /// <param name="initialRow">Row to start copying data from</param>
        public static void Create(
            SpreadsheetDocument document,
            List<string> headerList,
            string[][] valueTable,
            int initialRow
        )
        {
            WorksheetAccessor.Create(document, headerList, valueTable, initialRow);
        }
    }
}
