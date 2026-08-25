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
    private sealed class SlideWriter
    {
        private readonly SvgDocument _document;
        private readonly long _slideWidth;
        private readonly long _slideHeight;
        private readonly Dictionary<string, uint> _sourceIds = new(StringComparer.Ordinal);
        private readonly Dictionary<RenderElement, XElement> _connectorXml = new();
        private uint _nextId;

        public SlideWriter(SvgDocument document, uint nextId, long slideWidth, long slideHeight)
        {
            _document = document;
            _slideWidth = slideWidth;
            _slideHeight = slideHeight;
            _nextId = nextId;
        }

        public uint NextShapeId => _nextId;

        public void Write(XElement shapeTree)
        {
            foreach (var element in _document.Elements)
            {
                var id = _nextId++;
                if (!string.IsNullOrWhiteSpace(element.SourceId))
                    _sourceIds[element.SourceId!] = id;
                if (element.Kind == RenderKind.Shape)
                    shapeTree.Add(CreateShape(element, id));
                else if (element.Kind == RenderKind.Text)
                    shapeTree.Add(CreateText(element, id));
                else if (element.Kind == RenderKind.Path)
                    shapeTree.Add(CreatePath(element, id));
                else
                {
                    var connector = CreateConnector(element, id);
                    _connectorXml[element] = connector;
                    shapeTree.Add(connector);
                }
            }
            foreach (var alias in _document.Aliases)
            {
                if (_sourceIds.TryGetValue(alias.Value, out var childId))
                    _sourceIds[alias.Key] = childId;
            }
            foreach (
                var element in _document.Elements.Where(e =>
                    e.Kind == RenderKind.Connector && e.From is not null && e.To is not null
                )
            )
            {
                // Connection references are inserted after all source IDs are known.
                if (!_connectorXml.TryGetValue(element, out var connector))
                    continue;
                if (!_sourceIds.TryGetValue(element.From!.Id, out var fromId))
                    throw new FormatException($"Connector source '{element.From.Id}' was not found.");
                if (!_sourceIds.TryGetValue(element.To!.Id, out var toId))
                    throw new FormatException($"Connector destination '{element.To.Id}' was not found.");
                var nv = connector.Element(Presentation.nvCxnSpPr)!;
                var cNv = nv.Element(Presentation.cNvCxnSpPr)!;
                cNv.Add(
                    new XElement(
                        Drawing.stCxn,
                        new XAttribute(NoNamespace.id, fromId),
                        new XAttribute(NoNamespace.idx, AnchorIndex(element.From.Anchor))
                    )
                );
                cNv.Add(
                    new XElement(
                        Drawing.endCxn,
                        new XAttribute(NoNamespace.id, toId),
                        new XAttribute(NoNamespace.idx, AnchorIndex(element.To.Anchor))
                    )
                );
            }
        }

        private XElement CreateShape(RenderElement e, uint id) =>
            new(
                Presentation.sp,
                new XElement(
                    Presentation.nvSpPr,
                    new XElement(
                        Presentation.cNvPr,
                        new XAttribute(NoNamespace.id, id),
                        new XAttribute(NoNamespace.name, e.SourceId ?? $"Shape {id}")
                    ),
                    new XElement(Presentation.cNvSpPr),
                    new XElement(Presentation.nvPr)
                ),
                CreateShapeProperties(e),
                e.Text is null ? null : CreateTextBody(e)
            );

        private XElement CreateText(RenderElement e, uint id) =>
            new(
                Presentation.sp,
                new XElement(
                    Presentation.nvSpPr,
                    new XElement(
                        Presentation.cNvPr,
                        new XAttribute(NoNamespace.id, id),
                        new XAttribute(NoNamespace.name, e.SourceId ?? $"Text {id}")
                    ),
                    new XElement(Presentation.cNvSpPr, new XAttribute(NoNamespace.txBox, 1)),
                    new XElement(Presentation.nvPr)
                ),
                new XElement(
                    Presentation.spPr,
                    OffExt(e),
                    new XElement(Drawing.noFill),
                    new XElement(
                        Drawing.ln,
                        new XElement(Drawing.noFill),
                        new XElement(Drawing.headEnd),
                        new XElement(Drawing.tailEnd)
                    )
                ),
                CreateTextBody(e)
            );

        private XElement CreatePath(RenderElement e, uint id) =>
            new(
                Presentation.sp,
                new XElement(
                    Presentation.nvSpPr,
                    new XElement(
                        Presentation.cNvPr,
                        new XAttribute(NoNamespace.id, id),
                        new XAttribute(NoNamespace.name, e.SourceId ?? $"Path {id}")
                    ),
                    new XElement(Presentation.cNvSpPr),
                    new XElement(Presentation.nvPr)
                ),
                new XElement(
                    Presentation.spPr,
                    OffExt(e),
                    CreatePathGeometry(e),
                    Fill(e.Style),
                    CreateLine(e.Style, e.MarkerStart, e.MarkerEnd)
                )
            );

        private static XElement CreatePathGeometry(RenderElement e)
        {
            var width = Math.Max(1, (int)Math.Round(e.Width * 1000));
            var height = Math.Max(1, (int)Math.Round(e.Height * 1000));
            var path = new XElement(
                Drawing.path,
                new XAttribute(NoNamespace.w, width),
                new XAttribute(NoNamespace.h, height),
                e.Path!.Select(segment =>
                {
                    var points = segment.Points.Select(point => new XElement(
                        Drawing.pt,
                        new XAttribute(NoNamespace.x, GeometryCoordinate(point.X, e.X, e.Width, width)),
                        new XAttribute(NoNamespace.y, GeometryCoordinate(point.Y, e.Y, e.Height, height))
                    ));
                    return segment.Verb switch
                    {
                        PathVerb.Move => new XElement(Drawing.moveTo, points),
                        PathVerb.Line => new XElement(Drawing.lnTo, points),
                        PathVerb.Cubic => new XElement(Drawing.cubicBezTo, points),
                        PathVerb.Quadratic => new XElement(Drawing.quadBezTo, points),
                        PathVerb.Close => new XElement(Drawing.close),
                        _ => throw new InvalidDataException("Unknown SVG path segment."),
                    };
                })
            );
            return new XElement(
                Drawing.custGeom,
                new XElement(Drawing.avLst),
                new XElement(
                    Drawing.rect,
                    new XAttribute(NoNamespace.l, "l"),
                    new XAttribute(NoNamespace.t, "t"),
                    new XAttribute(NoNamespace.r, "r"),
                    new XAttribute(NoNamespace.b, "b")
                ),
                new XElement(Drawing.pathLst, path)
            );
        }

        private static int GeometryCoordinate(double value, double origin, double extent, int geometryExtent) =>
            (int)Math.Clamp(Math.Round((value - origin) / extent * geometryExtent), 0, geometryExtent);

        private XElement CreateConnector(RenderElement e, uint id) =>
            new(
                Presentation.cxnSp,
                new XElement(
                    Presentation.nvCxnSpPr,
                    new XElement(
                        Presentation.cNvPr,
                        new XAttribute(NoNamespace.id, id),
                        new XAttribute(NoNamespace.name, e.SourceId ?? $"Connector {id}")
                    ),
                    new XElement(Presentation.cNvCxnSpPr),
                    new XElement(Presentation.nvPr)
                ),
                new XElement(
                    Presentation.spPr,
                    ConnectorOffExt(e),
                    new XElement(
                        Drawing.prstGeom,
                        new XAttribute(NoNamespace.prst, "line"),
                        new XElement(Drawing.avLst)
                    ),
                    new XElement(Drawing.noFill),
                    CreateLine(e.Style, e.MarkerStart, e.MarkerEnd)
                )
            );

        private XElement CreateShapeProperties(RenderElement e) =>
            new XElement(
                Presentation.spPr,
                OffExt(e),
                new XElement(
                    Drawing.prstGeom,
                    new XAttribute(NoNamespace.prst, e.Geometry ?? "rect"),
                    new XElement(Drawing.avLst)
                ),
                Fill(e.Style),
                CreateLine(e.Style, e.MarkerStart, e.MarkerEnd)
            );

        private XElement OffExt(RenderElement e) =>
            new(
                Drawing.xfrm,
                e.Rotation == 0 ? null : new XAttribute(NoNamespace.rot, (long)Math.Round(e.Rotation * 60000)),
                new XElement(
                    Drawing.off,
                    new XAttribute(NoNamespace.x, EmuX(e.X)),
                    new XAttribute(NoNamespace.y, EmuY(e.Y))
                ),
                new XElement(
                    Drawing.ext,
                    new XAttribute(NoNamespace.cx, EmuX(e.Width)),
                    new XAttribute(NoNamespace.cy, EmuY(e.Height))
                )
            );

        private XElement ConnectorOffExt(RenderElement e)
        {
            var endX = e.X + e.Width;
            var endY = e.Y + e.Height;
            var x = Math.Min(e.X, endX);
            var y = Math.Min(e.Y, endY);
            var width = Math.Abs(e.Width);
            var height = Math.Abs(e.Height);
            return new XElement(
                Drawing.xfrm,
                e.Width < 0 ? new XAttribute(NoNamespace.flipH, 1) : null,
                e.Height < 0 ? new XAttribute(NoNamespace.flipV, 1) : null,
                new XElement(
                    Drawing.off,
                    new XAttribute(NoNamespace.x, EmuX(x)),
                    new XAttribute(NoNamespace.y, EmuY(y))
                ),
                new XElement(
                    Drawing.ext,
                    new XAttribute(NoNamespace.cx, EmuX(width)),
                    new XAttribute(NoNamespace.cy, EmuY(height))
                )
            );
        }

        private static XElement Fill(Style style) =>
            style.Fill is null or "none"
                ? new XElement(Drawing.noFill)
                : new XElement(Drawing.solidFill, Color(style.Fill));

        private XElement CreateLine(Style style, string? markerStart = null, string? markerEnd = null)
        {
            var line = new XElement(
                Drawing.ln,
                new XAttribute(NoNamespace.w, Math.Max(1, EmuX(style.StrokeWidth))),
                style.Stroke is null or "none"
                    ? new XElement(Drawing.noFill)
                    : new XElement(Drawing.solidFill, Color(style.Stroke)),
                Arrow("headEnd", markerStart),
                Arrow("tailEnd", markerEnd)
            );
            return line;
        }

        private static XElement Arrow(string name, string? marker)
        {
            var arrow = new XElement(Drawing.a + name);
            if (!string.IsNullOrWhiteSpace(marker))
                arrow.SetAttributeValue("type", "triangle");
            return arrow;
        }

        private static XElement CreateTextBody(RenderElement e) =>
            new XElement(
                Presentation.txBody,
                new XElement(
                    Drawing.bodyPr,
                    new XAttribute(NoNamespace.wrap, "square"),
                    new XAttribute(NoNamespace.lIns, 0),
                    new XAttribute(NoNamespace.rIns, 0),
                    new XAttribute(NoNamespace.tIns, 0),
                    new XAttribute(NoNamespace.bIns, 0),
                    new XAttribute(NoNamespace.anchor, "ctr")
                ),
                new XElement(Drawing.lstStyle),
                new XElement(
                    Drawing.p,
                    new XElement(
                        Drawing.pPr,
                        new XAttribute(
                            "algn",
                            e.Style.TextAnchor == "middle" ? "ctr"
                                : e.Style.TextAnchor == "end" ? "r"
                                : "l"
                        )
                    ),
                    new XElement(Drawing.r, RunProperties(e.Style), new XElement(Drawing.t, e.Text)),
                    new XElement(Drawing.endParaRPr, new XAttribute(NoNamespace.sz, Points(e.Style.FontSize)))
                )
            );

        private static XElement RunProperties(Style style) =>
            new XElement(
                Drawing.rPr,
                new XAttribute(NoNamespace.lang, "en-US"),
                new XAttribute(NoNamespace.sz, Points(style.FontSize)),
                new XAttribute(NoNamespace.b, style.FontWeight == "bold" ? 1 : 0),
                new XAttribute(NoNamespace.i, style.FontStyle == "italic" ? 1 : 0),
                ColorFill(style.Color),
                new XElement(Drawing.latin, new XAttribute(NoNamespace.typeface, style.FontFamily))
            );

        private static XElement ColorFill(string color) => new XElement(Drawing.solidFill, Color(color));

        private static XElement Color(string color) =>
            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, SvgColorParser.Normalize(color)));

        private long EmuX(double value) => (long)Math.Round(value * _slideWidth / 1280d);

        private long EmuY(double value) => (long)Math.Round(value * _slideHeight / 720d);

        private static int Points(double value) => Math.Max(1, (int)Math.Round(value * 100 * 0.75));

        private static int AnchorIndex(string anchor) =>
            anchor.ToLowerInvariant() switch
            {
                "top" => 0,
                "right" => 1,
                "bottom" => 2,
                "left" => 3,
                _ => 0,
            };
    }
}
