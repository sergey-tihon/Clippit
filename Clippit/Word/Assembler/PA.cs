using System.Xml.Linq;

namespace Clippit.Word.Assembler;

internal static class PA
{
    public static readonly XName Image = "Image";
    public static readonly XName Content = "Content";
    public static readonly XName DocumentTemplate = "DocumentTemplate";
    public static readonly XName Document = "Document";
    public static readonly XName Table = "Table";
    public static readonly XName Repeat = "Repeat";
    public static readonly XName EndRepeat = "EndRepeat";
    public static readonly XName Conditional = "Conditional";
    public static readonly XName EndConditional = "EndConditional";

    public static readonly XName Select = "Select";
    public static readonly XName Optional = "Optional";
    public static readonly XName Match = "Match";
    public static readonly XName NotMatch = "NotMatch";
    public static readonly XName Depth = "Depth";
    public static readonly XName Align = "Align";
    public static readonly XName Path = "Path";
    public static readonly XName Data = "Data";
    public static readonly XName PageBreakAfter = "PageBreakAfter";
    public static readonly XName FitWithin = "FitWithin";
}
