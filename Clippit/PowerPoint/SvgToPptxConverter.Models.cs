using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using Clippit.PowerPoint.Fluent;
using DocumentFormat.OpenXml.Packaging;
using Drawing = Clippit.A;
using Presentation = Clippit.P;
using Relationships = Clippit.R;

namespace Clippit.PowerPoint;

public static partial class SvgToPptxConverter
{
    private sealed record SvgDocument(
        double ScaleX,
        double ScaleY,
        IReadOnlyList<RenderElement> Elements,
        IReadOnlyDictionary<string, string> Aliases
    );

    private sealed class RenderElement
    {
        public RenderKind Kind { get; init; }
        public string? SourceId { get; set; }
        public double X { get; init; }
        public double Y { get; init; }
        public double Width { get; init; }
        public double Height { get; init; }
        public double Rotation { get; init; }
        public string? Geometry { get; init; }
        public string? Text { get; init; }
        public IReadOnlyList<PathSegment>? Path { get; init; }
        public string? MarkerStart { get; init; }
        public string? MarkerEnd { get; init; }
        public Style Style { get; init; } = Style.Default;
        public Endpoint? From { get; init; }
        public Endpoint? To { get; init; }
    }

    private enum RenderKind
    {
        Shape,
        Text,
        Path,
        Connector,
    }

    private enum PathVerb
    {
        Move,
        Line,
        Cubic,
        Quadratic,
        Close,
    }

    private sealed record PathSegment(PathVerb Verb, IReadOnlyList<Point> Points);

    private sealed record Endpoint(string Id, string Anchor);

    private readonly record struct Box(double X, double Y, double Width, double Height);

    private readonly record struct Point(double X, double Y);

    private readonly record struct Matrix(double A, double B, double C, double D, double E, double F)
    {
        public static Matrix Identity => new(1, 0, 0, 1, 0, 0);

        public Point Apply(double x, double y) => new(A * x + C * y + E, B * x + D * y + F);

        public Matrix Multiply(Matrix other) =>
            new(
                A * other.A + C * other.B,
                B * other.A + D * other.B,
                A * other.C + C * other.D,
                B * other.C + D * other.D,
                A * other.E + C * other.F + E,
                B * other.E + D * other.F + F
            );
    }

    private sealed record Style(
        string Fill,
        string Stroke,
        double StrokeWidth,
        string Color,
        string FontFamily,
        double FontSize,
        string FontWeight,
        string FontStyle,
        string TextAnchor,
        string Display,
        string Visibility,
        string? MarkerStart,
        string? MarkerEnd
    )
    {
        public static Style Default =>
            new("none", "none", 1, "000000", "Aptos", 18, "normal", "normal", "start", "inline", "visible", null, null);

        public Style Merge(XElement element)
        {
            var style = this;
            var inline = (string?)element.Attribute("style");
            var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (
                var pair in new[]
                {
                    ("fill", Fill),
                    ("stroke", Stroke),
                    ("stroke-width", StrokeWidth.ToString(CultureInfo.InvariantCulture)),
                    ("color", Color),
                    ("font-family", FontFamily),
                    ("font-size", FontSize.ToString(CultureInfo.InvariantCulture)),
                    ("font-weight", FontWeight),
                    ("font-style", FontStyle),
                    ("text-anchor", TextAnchor),
                    ("display", Display),
                    ("visibility", Visibility),
                    ("marker-start", MarkerStart ?? string.Empty),
                    ("marker-end", MarkerEnd ?? string.Empty),
                }
            )
                values[pair.Item1] = pair.Item2;
            foreach (var part in (inline ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries))
            {
                var p = part.Split(':', 2);
                if (p.Length == 2)
                    values[p[0].Trim()] = p[1].Trim();
            }
            foreach (var name in values.Keys.ToArray())
                if (element.Attribute(name) is { } a)
                    values[name] = a.Value;
            return style with
            {
                Fill = values["fill"],
                Stroke = values["stroke"],
                StrokeWidth = ParseDouble(values["stroke-width"], StrokeWidth),
                Color = values["color"],
                FontFamily = values["font-family"],
                FontSize = ParseDouble(values["font-size"], FontSize),
                FontWeight = values["font-weight"],
                FontStyle = values["font-style"],
                TextAnchor = values["text-anchor"],
                Display = values["display"],
                Visibility = values["visibility"],
                MarkerStart = string.IsNullOrWhiteSpace(values["marker-start"]) ? null : values["marker-start"],
                MarkerEnd = string.IsNullOrWhiteSpace(values["marker-end"]) ? null : values["marker-end"],
            };
        }

        private static double ParseDouble(string value, double fallback) =>
            double.TryParse(
                value.TrimEnd('p', 'x', 't'),
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var number
            )
                ? number
                : fallback;
    }

    private static Matrix ParseTransform(string? value)
    {
        var result = Matrix.Identity;
        if (string.IsNullOrWhiteSpace(value))
            return result;
        foreach (
            var match in System
                .Text.RegularExpressions.Regex.Matches(
                    value,
                    @"(?<name>translate|scale|rotate)\s*\((?<args>[^)]*)\)",
                    System.Text.RegularExpressions.RegexOptions.CultureInvariant
                )
                .Cast<System.Text.RegularExpressions.Match>()
        )
        {
            var name = match.Groups["name"].Value;
            var args = match
                .Groups["args"]
                .Value.Split([',', ' ', '\t'], StringSplitOptions.RemoveEmptyEntries)
                .Select(v => double.Parse(v, CultureInfo.InvariantCulture))
                .ToArray();
            var transform = name switch
            {
                "translate" => new Matrix(1, 0, 0, 1, args[0], args.Length > 1 ? args[1] : 0),
                "scale" => new Matrix(args[0], 0, 0, args.Length > 1 ? args[1] : args[0], 0, 0),
                "rotate" => Rotation(args[0], args.Length > 2 ? args[1] : 0, args.Length > 2 ? args[2] : 0),
                _ => Matrix.Identity,
            };
            result = result.Multiply(transform);
        }
        return result;
    }

    private static Matrix Rotation(double degrees, double cx, double cy)
    {
        var radians = degrees * Math.PI / 180;
        var cos = Math.Cos(radians);
        var sin = Math.Sin(radians);
        return new Matrix(cos, sin, -sin, cos, cx - cos * cx + sin * cy, cy - sin * cx - cos * cy);
    }
}

/// <summary>One SVG document used as a slide source.</summary>
public sealed record SvgSlideSource(string Content, string? Name = null);

/// <summary>Safety and output options for SVG conversion.</summary>
public sealed record SvgToPptxConverterSettings
{
    /// <summary>
    /// Existing presentation used as the output package. Generated SVG objects are appended to its slides.
    /// </summary>
    public PmlDocument? Template { get; init; }

    /// <summary>
    /// Zero-based slide index used for the first SVG. Subsequent SVGs use subsequent slides.
    /// </summary>
    public int TemplateSlideIndex { get; init; }

    public bool ValidateOutput { get; init; } = true;
    public int MaximumElements { get; init; } = 10_000;
    public long MaximumCharacters { get; init; } = 10_000_000;
}

internal static class SvgColorParser
{
    public static string Normalize(string value)
    {
        value = value.Trim();
        if (value.StartsWith('#'))
        {
            var hex = value[1..];
            if (hex.Length == 3)
                hex = string.Concat(hex.Select(c => $"{c}{c}"));
            if (hex.Length == 6 && hex.All(Uri.IsHexDigit))
                return hex.ToUpperInvariant();
        }
        if (value.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase))
        {
            var values = value[4..^1]
                .Split(',')
                .Select(v => int.Parse(v.Trim(), CultureInfo.InvariantCulture))
                .ToArray();
            if (values.Length == 3)
                return string.Concat(values.Select(v => Math.Clamp(v, 0, 255).ToString("X2")));
        }
        return value.ToLowerInvariant() switch
        {
            "white" => "FFFFFF",
            "black" => "000000",
            "red" => "FF0000",
            "green" => "008000",
            "blue" => "0000FF",
            "yellow" => "FFFF00",
            _ => "000000",
        };
    }
}
