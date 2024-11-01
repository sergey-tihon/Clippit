using System.Xml.Linq;

namespace Clippit.Word.Assembler
{
    internal static class PA
    {
        public static XName Image = "Image";
        public static XName Content = "Content";
        public static XName DocumentTemplate = "DocumentTemplate";
        public static XName Document = "Document";
        public static XName Table = "Table";
        public static XName Repeat = "Repeat";
        public static XName EndRepeat = "EndRepeat";
        public static XName Conditional = "Conditional";
        public static XName EndConditional = "EndConditional";

        public static XName Select = "Select";
        public static XName Optional = "Optional";
        public static XName Match = "Match";
        public static XName NotMatch = "NotMatch";
        public static XName Depth = "Depth";
        public static XName Align = "Align";
        public static XName Path = "Path";
        public static XName Data = "Data";
        public static XName PageBreakAfter = "PageBreakAfter";
    }
}
