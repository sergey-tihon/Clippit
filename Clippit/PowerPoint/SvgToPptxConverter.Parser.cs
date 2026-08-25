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
    private sealed class SvgParser
    {
        private readonly SvgToPptxConverterSettings _settings;
        private readonly List<RenderElement> _elements = [];
        private readonly Dictionary<string, RenderElement> _ids = new(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _aliases = new(StringComparer.Ordinal);
        private int _elementCount;
        private readonly string? _name;
        private double _scaleX;
        private double _scaleY;
        private double _viewMinX;
        private double _viewMinY;

        private SvgParser(string content, string? name, SvgToPptxConverterSettings settings)
        {
            _settings = settings;
            _name = name;
            var readerSettings = new XmlReaderSettings
            {
                // SVG 1.1 files commonly include a standard external DOCTYPE declaration. Ignore the
                // declaration rather than processing it so those files remain readable without enabling
                // external entity resolution.
                DtdProcessing = DtdProcessing.Ignore,
                XmlResolver = null,
                MaxCharactersInDocument = settings.MaximumCharacters,
                MaxCharactersFromEntities = 0,
            };
            using var stringReader = new StringReader(content);
            using var reader = XmlReader.Create(stringReader, readerSettings);
            var document = XDocument.Load(reader, LoadOptions.SetLineInfo);
            ParseRoot(document.Root ?? throw Error("SVG root is missing."));
        }

        public static SvgDocument Parse(string content, string? name, SvgToPptxConverterSettings settings) =>
            new SvgParser(content, name, settings).CreateDocument();

        private SvgDocument CreateDocument() => new(_scaleX, _scaleY, _elements, _aliases);

        private void ParseRoot(XElement root)
        {
            if (root.Name != S + "svg")
                throw Error("The document root must be an SVG element.");
            if (root.Descendants().Any(e => e.Name.LocalName is "script" or "iframe"))
                throw Error("Scripts and iframe content are not supported.");
            var viewBox = ParseNumbers(root.Attribute("viewBox")?.Value);
            if (
                viewBox.Length != 4
                || viewBox.Any(double.IsNaN)
                || viewBox.Any(double.IsInfinity)
                || viewBox[2] <= 0
                || viewBox[3] <= 0
            )
                throw Error("SVG must contain a finite, positive four-value viewBox.");
            _viewMinX = viewBox[0];
            _viewMinY = viewBox[1];
            _scaleX = 1280d / viewBox[2];
            _scaleY = 720d / viewBox[3];
            Visit(root, new Matrix(_scaleX, 0, 0, _scaleY, -_viewMinX * _scaleX, -_viewMinY * _scaleY), Style.Default);
        }

        private void Visit(XElement element, Matrix parentTransform, Style parentStyle)
        {
            if (
                element.Name.LocalName
                is "defs"
                    or "style"
                    or "marker"
                    or "clipPath"
                    or "mask"
                    or "filter"
                    or "foreignObject"
            )
                return;

            if (++_elementCount > _settings.MaximumElements)
                throw Error($"SVG exceeds the maximum element count of {_settings.MaximumElements}.");
            var style = parentStyle.Merge(element);
            var transform = parentTransform.Multiply(ParseTransform(element.Attribute("transform")?.Value));
            var id = (string?)element.Attribute("id");
            var firstChildIndex = _elements.Count;
            RenderElement? created = element.Name.LocalName switch
            {
                "rect" => CreateRect(element, transform, style),
                "circle" => CreateCircle(element, transform, style),
                "ellipse" => CreateEllipse(element, transform, style),
                "line" => CreateLine(element, transform, style),
                "path" => CreatePath(element, transform, style),
                "text" => CreateText(element, transform, style),
                _ => null,
            };
            if (created is not null)
            {
                _elements.Add(created);
                if (!string.IsNullOrWhiteSpace(id))
                {
                    if (!_ids.TryAdd(id, created))
                        throw Error($"Duplicate SVG id '{id}'.");
                }
            }
            foreach (var child in element.Elements())
                Visit(child, transform, style);
            if (id is not null && created is null)
            {
                var child = _elements.Skip(firstChildIndex).FirstOrDefault(e => e.Kind == RenderKind.Shape);
                if (child is not null)
                {
                    child.SourceId ??= id;
                    _aliases[id] = child.SourceId;
                }
            }
        }

        private RenderElement? CreateRect(XElement e, Matrix t, Style style)
        {
            var x = Number(e, "x");
            var y = Number(e, "y");
            var width = Number(e, "width");
            var height = Number(e, "height");
            if (width <= 0 || height <= 0 || style.Display == "none" || style.Visibility == "hidden")
                return null;
            return Shape(
                e,
                t,
                style,
                new Box(x, y, width, height),
                (Number(e, "rx") > 0 || Number(e, "ry") > 0) ? "roundRect" : "rect"
            );
        }

        private RenderElement? CreateCircle(XElement e, Matrix t, Style style)
        {
            var r = Number(e, "r");
            if (r <= 0)
                return null;
            return Shape(e, t, style, new Box(Number(e, "cx") - r, Number(e, "cy") - r, r * 2, r * 2), "ellipse");
        }

        private RenderElement? CreateEllipse(XElement e, Matrix t, Style style)
        {
            var rx = Number(e, "rx");
            var ry = Number(e, "ry");
            if (rx <= 0 || ry <= 0)
                return null;
            return Shape(e, t, style, new Box(Number(e, "cx") - rx, Number(e, "cy") - ry, rx * 2, ry * 2), "ellipse");
        }

        private RenderElement? CreateLine(XElement e, Matrix t, Style style)
        {
            if (style.Display == "none" || style.Visibility == "hidden")
                return null;
            var start = t.Apply(Number(e, "x1"), Number(e, "y1"));
            var end = t.Apply(Number(e, "x2"), Number(e, "y2"));
            var connector = new RenderElement
            {
                Kind = RenderKind.Connector,
                SourceId = (string?)e.Attribute("id"),
                X = start.X,
                Y = start.Y,
                Width = end.X - start.X,
                Height = end.Y - start.Y,
                Style = style,
                MarkerStart = style.MarkerStart,
                MarkerEnd = style.MarkerEnd,
                From = ParseEndpoint((string?)e.Attribute("data-pptx-from")),
                To = ParseEndpoint((string?)e.Attribute("data-pptx-to")),
            };
            return connector;
        }

        private RenderElement? CreatePath(XElement e, Matrix t, Style style)
        {
            if (style.Display == "none" || style.Visibility == "hidden")
                return null;

            var path = ParsePath((string?)e.Attribute("d"), t);
            if (path.Count == 0)
                return null;

            var points = path.SelectMany(segment => segment.Points).ToArray();
            if (points.Length == 0)
                return null;

            var minX = points.Min(point => point.X);
            var minY = points.Min(point => point.Y);
            var maxX = points.Max(point => point.X);
            var maxY = points.Max(point => point.Y);
            return new RenderElement
            {
                Kind = RenderKind.Path,
                SourceId = (string?)e.Attribute("id"),
                X = minX,
                Y = minY,
                Width = Math.Max(maxX - minX, 0.001),
                Height = Math.Max(maxY - minY, 0.001),
                Style = style,
                Path = path,
                MarkerStart = style.MarkerStart,
                MarkerEnd = style.MarkerEnd,
            };
        }

        private static IReadOnlyList<PathSegment> ParsePath(string? data, Matrix transform)
        {
            if (string.IsNullOrWhiteSpace(data))
                return [];

            var tokens = System
                .Text.RegularExpressions.Regex.Matches(
                    data,
                    @"[AaCcHhLlMmQqSsTtVvZz]|[-+]?(?:(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?)",
                    System.Text.RegularExpressions.RegexOptions.CultureInvariant
                )
                .Select(match => match.Value)
                .ToArray();
            var segments = new List<PathSegment>();
            var index = 0;
            var command = '\0';
            var sourceCurrent = new Point(0, 0);
            var sourceSubpathStart = sourceCurrent;
            var previousCubicControl = sourceCurrent;
            var previousQuadraticControl = sourceCurrent;
            var previousCommand = '\0';

            while (index < tokens.Length)
            {
                if (IsPathCommand(tokens[index]))
                    command = tokens[index++][0];
                else if (command == '\0')
                    throw new FormatException("Invalid SVG path: a command is required.");

                var relative = char.IsLower(command);
                var verb = char.ToUpperInvariant(command);
                if (verb == 'Z')
                {
                    segments.Add(new PathSegment(PathVerb.Close, []));
                    sourceCurrent = sourceSubpathStart;
                    previousCommand = command;
                    command = '\0';
                    continue;
                }

                var required = verb switch
                {
                    'M' or 'L' or 'T' => 2,
                    'H' or 'V' => 1,
                    'C' => 6,
                    'S' or 'Q' => 4,
                    _ => throw new FormatException($"SVG path command '{command}' is not supported."),
                };
                if (index + required > tokens.Length || IsPathCommand(tokens[index]))
                    throw new FormatException($"SVG path command '{command}' has incomplete arguments.");
                var values = new double[required];
                for (var valueIndex = 0; valueIndex < required; valueIndex++)
                {
                    if (
                        !double.TryParse(
                            tokens[index++],
                            NumberStyles.Float,
                            CultureInfo.InvariantCulture,
                            out values[valueIndex]
                        )
                    )
                        throw new FormatException("SVG path contains an invalid number.");
                }

                Point PointAt(double x, double y) =>
                    relative ? new Point(sourceCurrent.X + x, sourceCurrent.Y + y) : new Point(x, y);
                switch (verb)
                {
                    case 'M':
                    {
                        sourceCurrent = PointAt(values[0], values[1]);
                        sourceSubpathStart = sourceCurrent;
                        var transformedMove = transform.Apply(sourceCurrent.X, sourceCurrent.Y);
                        segments.Add(new PathSegment(PathVerb.Move, [transformedMove]));
                        command = relative ? 'l' : 'L';
                        previousCommand = verb;
                        break;
                    }
                    case 'L':
                    {
                        sourceCurrent = PointAt(values[0], values[1]);
                        var transformedLine = transform.Apply(sourceCurrent.X, sourceCurrent.Y);
                        segments.Add(new PathSegment(PathVerb.Line, [transformedLine]));
                        previousCommand = verb;
                        break;
                    }
                    case 'H':
                    {
                        sourceCurrent = relative
                            ? new Point(sourceCurrent.X + values[0], sourceCurrent.Y)
                            : new Point(values[0], sourceCurrent.Y);
                        var transformedHorizontal = transform.Apply(sourceCurrent.X, sourceCurrent.Y);
                        segments.Add(new PathSegment(PathVerb.Line, [transformedHorizontal]));
                        previousCommand = verb;
                        break;
                    }
                    case 'V':
                    {
                        sourceCurrent = relative
                            ? new Point(sourceCurrent.X, sourceCurrent.Y + values[0])
                            : new Point(sourceCurrent.X, values[0]);
                        var transformedVertical = transform.Apply(sourceCurrent.X, sourceCurrent.Y);
                        segments.Add(new PathSegment(PathVerb.Line, [transformedVertical]));
                        previousCommand = verb;
                        break;
                    }
                    case 'C':
                    {
                        var control1 = PointAt(values[0], values[1]);
                        var control2 = PointAt(values[2], values[3]);
                        var endpoint = PointAt(values[4], values[5]);
                        sourceCurrent = endpoint;
                        previousCubicControl = control2;
                        var transformedEndpoint = transform.Apply(endpoint.X, endpoint.Y);
                        segments.Add(
                            new PathSegment(
                                PathVerb.Cubic,
                                [
                                    transform.Apply(control1.X, control1.Y),
                                    transform.Apply(control2.X, control2.Y),
                                    transformedEndpoint,
                                ]
                            )
                        );
                        previousCommand = verb;
                        break;
                    }
                    case 'S':
                    {
                        var control1 = previousCommand is 'C' or 'S'
                            ? new Point(
                                sourceCurrent.X * 2 - previousCubicControl.X,
                                sourceCurrent.Y * 2 - previousCubicControl.Y
                            )
                            : sourceCurrent;
                        var control2 = PointAt(values[0], values[1]);
                        var endpoint = PointAt(values[2], values[3]);
                        sourceCurrent = endpoint;
                        previousCubicControl = control2;
                        var transformedEndpoint = transform.Apply(endpoint.X, endpoint.Y);
                        segments.Add(
                            new PathSegment(
                                PathVerb.Cubic,
                                [
                                    transform.Apply(control1.X, control1.Y),
                                    transform.Apply(control2.X, control2.Y),
                                    transformedEndpoint,
                                ]
                            )
                        );
                        previousCommand = verb;
                        break;
                    }
                    case 'Q':
                    {
                        var control = PointAt(values[0], values[1]);
                        var endpoint = PointAt(values[2], values[3]);
                        sourceCurrent = endpoint;
                        previousQuadraticControl = control;
                        var transformedEndpoint = transform.Apply(endpoint.X, endpoint.Y);
                        segments.Add(
                            new PathSegment(
                                PathVerb.Quadratic,
                                [transform.Apply(control.X, control.Y), transformedEndpoint]
                            )
                        );
                        previousCommand = verb;
                        break;
                    }
                    case 'T':
                    {
                        var control = previousCommand is 'Q' or 'T'
                            ? new Point(
                                sourceCurrent.X * 2 - previousQuadraticControl.X,
                                sourceCurrent.Y * 2 - previousQuadraticControl.Y
                            )
                            : sourceCurrent;
                        var endpoint = PointAt(values[0], values[1]);
                        sourceCurrent = endpoint;
                        previousQuadraticControl = control;
                        var transformedEndpoint = transform.Apply(endpoint.X, endpoint.Y);
                        segments.Add(
                            new PathSegment(
                                PathVerb.Quadratic,
                                [transform.Apply(control.X, control.Y), transformedEndpoint]
                            )
                        );
                        previousCommand = verb;
                        break;
                    }
                }
            }
            return segments;
        }

        private static bool IsPathCommand(string token) => token.Length == 1 && char.IsLetter(token[0]);

        private RenderElement? CreateText(XElement e, Matrix t, Style style)
        {
            if (style.Display == "none" || style.Visibility == "hidden")
                return null;
            var text = string.Concat(
                e.Nodes()
                    .Select(n =>
                        n is XText tNode ? tNode.Value
                        : n is XElement x && x.Name == S + "tspan" ? x.Value
                        : string.Empty
                    )
            );
            if (string.IsNullOrWhiteSpace(text))
                return null;
            var point = t.Apply(Number(e, "x"), Number(e, "y"));
            var textScale = Math.Sqrt(t.A * t.A + t.B * t.B);
            var fontSize = style.FontSize * textScale;
            var width = Math.Max(fontSize * text.Length * 0.55, fontSize);
            var height = fontSize * 1.35;
            if (style.TextAnchor == "middle")
                point = new Point(point.X - width / 2, point.Y - height);
            else if (style.TextAnchor == "end")
                point = new Point(point.X - width, point.Y - height);
            else
                point = new Point(point.X, point.Y - height);
            return new RenderElement
            {
                Kind = RenderKind.Text,
                SourceId = (string?)e.Attribute("id"),
                X = point.X,
                Y = point.Y,
                Width = width,
                Height = height,
                Style = (style.Fill is not "none" ? style with { Color = style.Fill } : style) with
                {
                    FontSize = fontSize,
                },
                Text = text,
            };
        }

        private RenderElement Shape(XElement e, Matrix transform, Style style, Box box, string geometry)
        {
            var corners = new[]
            {
                transform.Apply(box.X, box.Y),
                transform.Apply(box.X + box.Width, box.Y),
                transform.Apply(box.X, box.Y + box.Height),
                transform.Apply(box.X + box.Width, box.Y + box.Height),
            };
            var scaleX = Math.Sqrt(transform.A * transform.A + transform.B * transform.B);
            var scaleY = Math.Sqrt(transform.C * transform.C + transform.D * transform.D);
            var center = transform.Apply(box.X + box.Width / 2, box.Y + box.Height / 2);
            var width = box.Width * scaleX;
            var height = box.Height * scaleY;
            var rotation = Math.Atan2(transform.B, transform.A) * 180 / Math.PI;
            return new RenderElement
            {
                Kind = RenderKind.Shape,
                SourceId = (string?)e.Attribute("id"),
                X = center.X - width / 2,
                Y = center.Y - height / 2,
                Width = width,
                Height = height,
                Rotation = rotation,
                Geometry = geometry,
                Style = style,
            };
        }

        private double Number(XElement e, string name) =>
            double.TryParse(e.Attribute(name)?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value)
                ? value
                : 0;

        private static double[] ParseNumbers(string? value) =>
            (value ?? string.Empty)
                .Split([',', ' ', '\t', '\r', '\n'], StringSplitOptions.RemoveEmptyEntries)
                .Select(v =>
                    double.TryParse(v, NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : double.NaN
                )
                .ToArray();

        private static Endpoint? ParseEndpoint(string? value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return null;
            var parts = value.Split(':', 2);
            return new Endpoint(parts[0], parts.Length == 2 ? parts[1] : "center");
        }

        private Exception Error(string message) =>
            new FormatException($"Invalid SVG{(_name is null ? string.Empty : $" '{_name}'")}: {message}");
    }
}
