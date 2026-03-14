namespace Clippit.Word.Assembler;

public static class StringExtensions
{
    public static IList<string> SplitAndKeep(this string s, params char[] delimiters)
    {
        var parts = new List<string>();
        if (string.IsNullOrEmpty(s))
            return parts;

        var first = 0;
        do
        {
            var last = s.IndexOfAny(delimiters, first);
            if (last >= 0)
            {
                if (last > first)
                    parts.Add(s[first..last]);

                parts.Add(s[last].ToString());
                first = last + 1;
                continue;
            }

            parts.Add(s[first..]);
            break;
        } while (first < s.Length);

        return parts;
    }
}
