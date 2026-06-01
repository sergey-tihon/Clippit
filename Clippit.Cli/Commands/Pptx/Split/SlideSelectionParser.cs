using System.Globalization;
using Clippit.Cli.Infrastructure;

namespace Clippit.Cli.Commands.Pptx.Split;

internal static class SlideSelectionParser
{
    public static List<int> Parse(string? expression, int slideCount)
    {
        if (string.IsNullOrWhiteSpace(expression))
            return Enumerable.Range(1, slideCount).ToList();

        var indexes = new List<int>();
        var seen = new HashSet<int>();
        var tokens = expression.Split(',');

        foreach (var rawToken in tokens)
        {
            var token = rawToken.Trim();
            if (token.Length == 0)
                throw Invalid(expression, "contains an empty segment");

            var rangeParts = token.Split('-', StringSplitOptions.TrimEntries);
            if (rangeParts.Length > 2 || rangeParts.Any(part => part.Length == 0))
                throw Invalid(expression, $"contains invalid segment '{token}'");

            if (!TryParseSlideNumber(rangeParts[0], out var start))
                throw Invalid(expression, $"contains invalid slide number '{rangeParts[0]}'");

            var end = start;
            if (rangeParts.Length == 2)
            {
                if (!TryParseSlideNumber(rangeParts[1], out end))
                    throw Invalid(expression, $"contains invalid slide number '{rangeParts[1]}'");

                if (end < start)
                    throw Invalid(expression, $"contains descending range '{token}'");
            }

            if (start > slideCount || end > slideCount)
                throw Invalid(
                    expression,
                    $"references slide {Math.Max(start, end)}, but the presentation has {slideCount} slides"
                );

            for (var i = start; i <= end; i++)
            {
                if (seen.Add(i))
                    indexes.Add(i);
            }
        }

        return indexes;
    }

    private static bool TryParseSlideNumber(string value, out int slideNumber) =>
        int.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out slideNumber) && slideNumber > 0;

    private static CliException Invalid(string expression, string reason) =>
        CliException.InvalidArguments($"Invalid slide selection '{expression}': {reason}.");
}
