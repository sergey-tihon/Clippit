using System;
using System.Collections.Generic;

namespace Clippit.Word.Assembler
{
    public static class StringExtensions
    {
        public static IList<string> SplitAndKeep(this string s, params char[] delimiters)
        {
            var parts = new List<string>();
            if (!string.IsNullOrEmpty(s))
            {
                int first = 0;
                do
                {
                    int last = s.IndexOfAny(delimiters, first);
                    if (last >= 0)
                    {
                        if (last > first)
                        {
                            parts.Add(s.Substring(first, last - first)); // part before the delimiter
                        }

                        parts.Add(new string(s[last], 1)); // the delimiter
                        first = last + 1;
                        continue;
                    }

                    // no delimiters were found, but at least one character remains. Add the rest and stop.
                    parts.Add(s.Substring(first, s.Length - first));
                    break;

                } while (first < s.Length);
            }

            return parts;
        }
    }
}
