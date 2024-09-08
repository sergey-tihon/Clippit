﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Linq;
using System.Text;

namespace Clippit.Excel
{
    public class ParseFormula
    {
        private readonly ExcelFormula _parser;

        public ParseFormula(string formula)
        {
            _parser = new ExcelFormula(formula, Console.Out);
            var parserResult = false;
            try
            {
                parserResult = _parser.Formula();
            }
            catch (PegException) { }
            if (!parserResult)
            {
                _parser.Warning("Error processing " + formula);
            }
        }

        public string ReplaceSheetName(string oldName, string newName)
        {
            var text = new StringBuilder(_parser.GetSource());
            ReplaceNode(_parser.GetRoot(), (int)EExcelFormula.SheetName, oldName, newName, text);
            return text.ToString();
        }

        public string ReplaceRelativeCell(int rowOffset, int colOffset)
        {
            var text = new StringBuilder(_parser.GetSource());
            ReplaceRelativeCell(_parser.GetRoot(), rowOffset, colOffset, text);
            return text.ToString();
        }

        // Recursive function that will replace values from last to first
        private void ReplaceNode(PegNode node, int id, string oldName, string newName, StringBuilder text)
        {
            if (node.next_ != null)
                ReplaceNode(node.next_, id, oldName, newName, text);
            if (node.id_ == id && _parser.GetSource().Substring(node.match_._posBeg, node.match_.Length) == oldName)
            {
                text.Remove(node.match_._posBeg, node.match_.Length);
                text.Insert(node.match_._posBeg, newName);
            }
            else if (node.child_ != null)
                ReplaceNode(node.child_, id, oldName, newName, text);
        }

        // Recursive function that will adjust relative cells from last to first
        private void ReplaceRelativeCell(PegNode node, int rowOffset, int colOffset, StringBuilder text)
        {
            if (node.next_ != null)
                ReplaceRelativeCell(node.next_, rowOffset, colOffset, text);
            if (node.id_ == (int)EExcelFormula.A1Row && _parser.GetSource().Substring(node.match_._posBeg, 1) != "$")
            {
                var rowNumber = Convert.ToInt32(_parser.GetSource().Substring(node.match_._posBeg, node.match_.Length));
                text.Remove(node.match_._posBeg, node.match_.Length);
                text.Insert(node.match_._posBeg, Convert.ToString(rowNumber + rowOffset));
            }
            else if (
                node.id_ == (int)EExcelFormula.A1Column
                && _parser.GetSource().Substring(node.match_._posBeg, 1) != "$"
            )
            {
                var colNumber = GetColumnNumber(_parser.GetSource().Substring(node.match_._posBeg, node.match_.Length));
                text.Remove(node.match_._posBeg, node.match_.Length);
                text.Insert(node.match_._posBeg, GetColumnId(colNumber + colOffset));
            }
            else if (node.child_ != null)
                ReplaceRelativeCell(node.child_, rowOffset, colOffset, text);
        }

        // Converts the column reference string to a column number (e.g. A -> 1, B -> 2)
        private static int GetColumnNumber(string cellReference) =>
            cellReference
                .Where(char.IsLetter)
                .Aggregate(0, (current, c) => current * 26 + Convert.ToInt32(c) - Convert.ToInt32('A') + 1);

        // Translates the column number to the column reference string (e.g. 1 -> A, 2-> B)
        private static string GetColumnId(int columnNumber)
        {
            var result = "";
            do
            {
                result = ((char)((columnNumber - 1) % 26 + 'A')) + result;
                columnNumber = (columnNumber - 1) / 26;
            } while (columnNumber != 0);
            return result;
        }
    }
}
