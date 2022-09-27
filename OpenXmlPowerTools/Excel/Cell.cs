using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace Clippit.Excel
{
    public static class Cell
    {
        public static CellDfn[] Headers(params string[] headers) =>
            headers.Select(value => String(value, true)).ToArray();

        public static CellDfn String(string value, bool bold = false) =>
            new() { CellDataType = CellDataType.String, Value = RemoveForbiddenChars(value), Bold = bold };

        public static CellDfn Number(int value) =>
            new() { CellDataType = CellDataType.Number, Value = value };

        public static CellDfn Number(long value) =>
            new() { CellDataType = CellDataType.Number, Value = value };

        public static CellDfn Bool(bool? value) =>
            new() { CellDataType = CellDataType.Boolean, Value = value };

        public static CellDfn Date(DateTime? value)
        {
            if (value is null || value.Value == DateTime.MinValue)
                return null;

            return new CellDfn { CellDataType = CellDataType.Date, Value = value.Value, FormatCode = "mm-dd-yy" };
        }

        // From xml spec valid chars:
        // #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
        // any Unicode character, excluding the surrogate blocks, FFFE, and FFFF.
        private static readonly Regex s_xmlInvalidSymbolsRegex =
            new(@"[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000-x10FFFF]", RegexOptions.Compiled);
        
        private static string RemoveForbiddenChars(string strInput)
        {
            if (string.IsNullOrWhiteSpace(strInput))
                return strInput;

            return s_xmlInvalidSymbolsRegex.Replace(strInput, string.Empty);
        }
    }
}
