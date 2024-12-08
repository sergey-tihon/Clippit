using System;

namespace Clippit.Word.Assembler
{
    internal static class UriExtensions
    {
        internal static Uri GetUri(this string s)
        {
            return new UriBuilder(s).Uri;
        }
    }
}
