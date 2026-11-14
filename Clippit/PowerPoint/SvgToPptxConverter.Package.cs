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
    private static void CreatePresentationParts(PresentationPart presentationPart)
    {
        // CreatePresentationDocument provides only an empty package shell; Clippit has no embedded PPTX resource.
        // Keep this as the single default-package factory. Template mode bypasses it and preserves the template's
        // existing master, layout, theme, relationships, and slide size.
        var master = presentationPart.AddNewPart<SlideMasterPart>();
        var layout = master.AddNewPart<SlideLayoutPart>();
        var theme = master.AddNewPart<ThemePart>();
        theme.PutXDocument(CreateThemeXml());
        layout.PutXDocument(CreateLayoutXml(master.GetIdOfPart(layout)));
        master.PutXDocument(CreateMasterXml(master.GetIdOfPart(layout)));
        layout.AddPart(master);
        var presentation = new XDocument(
            new XElement(
                Presentation.presentation,
                new XAttribute(XNamespace.Xmlns + "a", Drawing.a),
                new XAttribute(XNamespace.Xmlns + "p", Presentation.p),
                new XAttribute(XNamespace.Xmlns + "r", Relationships.r),
                new XElement(
                    Presentation.sldMasterIdLst,
                    new XElement(
                        Presentation.sldMasterId,
                        new XAttribute(NoNamespace.id, 2147483648u),
                        new XAttribute(Relationships.id, presentationPart.GetIdOfPart(master))
                    )
                ),
                new XElement(Presentation.sldIdLst),
                new XElement(
                    Presentation.sldSz,
                    new XAttribute(NoNamespace.cx, DefaultSlideWidth),
                    new XAttribute(NoNamespace.cy, DefaultSlideHeight)
                ),
                new XElement(
                    Presentation.notesSz,
                    new XAttribute(NoNamespace.cx, 6_858_000),
                    new XAttribute(NoNamespace.cy, 9_144_000)
                )
            )
        );
        presentationPart.PutXDocument(presentation);
    }

    private static XDocument CreateSlideXml() =>
        new(
            new XElement(
                Presentation.sld,
                new XAttribute(XNamespace.Xmlns + "a", Drawing.a),
                new XAttribute(XNamespace.Xmlns + "p", Presentation.p),
                new XAttribute(XNamespace.Xmlns + "r", Relationships.r),
                new XElement(
                    Presentation.cSld,
                    new XAttribute(NoNamespace.name, "SVG slide"),
                    new XElement(
                        Presentation.spTree,
                        new XElement(
                            Presentation.nvGrpSpPr,
                            new XElement(
                                Presentation.cNvPr,
                                new XAttribute(NoNamespace.id, 1),
                                new XAttribute(NoNamespace.name, "")
                            ),
                            new XElement(Presentation.cNvGrpSpPr),
                            new XElement(Presentation.nvPr)
                        ),
                        new XElement(
                            Presentation.grpSpPr,
                            new XElement(
                                Drawing.xfrm,
                                new XElement(
                                    Drawing.off,
                                    new XAttribute(NoNamespace.x, 0),
                                    new XAttribute(NoNamespace.y, 0)
                                ),
                                new XElement(
                                    Drawing.ext,
                                    new XAttribute(NoNamespace.cx, 0),
                                    new XAttribute(NoNamespace.cy, 0)
                                ),
                                new XElement(
                                    Drawing.chOff,
                                    new XAttribute(NoNamespace.x, 0),
                                    new XAttribute(NoNamespace.y, 0)
                                ),
                                new XElement(
                                    Drawing.chExt,
                                    new XAttribute(NoNamespace.cx, 0),
                                    new XAttribute(NoNamespace.cy, 0)
                                )
                            )
                        )
                    )
                ),
                new XElement(Presentation.clrMapOvr, new XElement(Drawing.masterClrMapping))
            )
        );

    private static XDocument CreateMasterXml(string layoutRelationshipId) =>
        new(
            new XElement(
                Presentation.sldMaster,
                new XAttribute(XNamespace.Xmlns + "a", Drawing.a),
                new XAttribute(XNamespace.Xmlns + "p", Presentation.p),
                new XAttribute(XNamespace.Xmlns + "r", Relationships.r),
                new XElement(
                    Presentation.cSld,
                    new XAttribute(NoNamespace.name, "Master"),
                    new XElement(
                        Presentation.spTree,
                        new XElement(
                            Presentation.nvGrpSpPr,
                            new XElement(
                                Presentation.cNvPr,
                                new XAttribute(NoNamespace.id, 1),
                                new XAttribute(NoNamespace.name, "")
                            ),
                            new XElement(Presentation.cNvGrpSpPr),
                            new XElement(Presentation.nvPr)
                        ),
                        new XElement(
                            Presentation.grpSpPr,
                            new XElement(
                                Drawing.xfrm,
                                new XElement(
                                    Drawing.off,
                                    new XAttribute(NoNamespace.x, 0),
                                    new XAttribute(NoNamespace.y, 0)
                                ),
                                new XElement(
                                    Drawing.ext,
                                    new XAttribute(NoNamespace.cx, 0),
                                    new XAttribute(NoNamespace.cy, 0)
                                ),
                                new XElement(
                                    Drawing.chOff,
                                    new XAttribute(NoNamespace.x, 0),
                                    new XAttribute(NoNamespace.y, 0)
                                ),
                                new XElement(
                                    Drawing.chExt,
                                    new XAttribute(NoNamespace.cx, 0),
                                    new XAttribute(NoNamespace.cy, 0)
                                )
                            )
                        )
                    )
                ),
                new XElement(
                    Presentation.clrMap,
                    new XAttribute(NoNamespace.bg1, "lt1"),
                    new XAttribute(NoNamespace.tx1, "dk1"),
                    new XAttribute(NoNamespace.bg2, "lt2"),
                    new XAttribute(NoNamespace.tx2, "dk2"),
                    new XAttribute(NoNamespace.accent1, "accent1"),
                    new XAttribute(NoNamespace.accent2, "accent2"),
                    new XAttribute(NoNamespace.accent3, "accent3"),
                    new XAttribute(NoNamespace.accent4, "accent4"),
                    new XAttribute(NoNamespace.accent5, "accent5"),
                    new XAttribute(NoNamespace.accent6, "accent6"),
                    new XAttribute(NoNamespace.hlink, "hlink"),
                    new XAttribute(NoNamespace.folHlink, "folHlink")
                ),
                new XElement(
                    Presentation.sldLayoutIdLst,
                    new XElement(
                        Presentation.sldLayoutId,
                        new XAttribute(NoNamespace.id, 2147483649u),
                        new XAttribute(Relationships.id, layoutRelationshipId)
                    )
                )
            )
        );

    private static XDocument CreateLayoutXml(string masterRelationshipId) =>
        new(
            new XElement(
                Presentation.sldLayout,
                new XAttribute(XNamespace.Xmlns + "a", Drawing.a),
                new XAttribute(XNamespace.Xmlns + "p", Presentation.p),
                new XAttribute(XNamespace.Xmlns + "r", Relationships.r),
                new XAttribute(NoNamespace.type, "blank"),
                new XAttribute(NoNamespace.preserve, "1"),
                new XElement(
                    Presentation.cSld,
                    new XAttribute(NoNamespace.name, "Blank"),
                    new XElement(
                        Presentation.spTree,
                        new XElement(
                            Presentation.nvGrpSpPr,
                            new XElement(
                                Presentation.cNvPr,
                                new XAttribute(NoNamespace.id, 1),
                                new XAttribute(NoNamespace.name, "")
                            ),
                            new XElement(Presentation.cNvGrpSpPr),
                            new XElement(Presentation.nvPr)
                        ),
                        new XElement(
                            Presentation.grpSpPr,
                            new XElement(
                                Drawing.xfrm,
                                new XElement(
                                    Drawing.off,
                                    new XAttribute(NoNamespace.x, 0),
                                    new XAttribute(NoNamespace.y, 0)
                                ),
                                new XElement(
                                    Drawing.ext,
                                    new XAttribute(NoNamespace.cx, 0),
                                    new XAttribute(NoNamespace.cy, 0)
                                ),
                                new XElement(
                                    Drawing.chOff,
                                    new XAttribute(NoNamespace.x, 0),
                                    new XAttribute(NoNamespace.y, 0)
                                ),
                                new XElement(
                                    Drawing.chExt,
                                    new XAttribute(NoNamespace.cx, 0),
                                    new XAttribute(NoNamespace.cy, 0)
                                )
                            )
                        )
                    )
                ),
                new XElement(Presentation.clrMapOvr, new XElement(Drawing.masterClrMapping))
            )
        );

    private static XDocument CreateThemeXml() =>
        new(
            new XElement(
                Drawing.theme,
                new XAttribute(NoNamespace.name, "Clippit"),
                new XElement(
                    Drawing.themeElements,
                    new XElement(
                        Drawing.clrScheme,
                        new XAttribute(NoNamespace.name, "Clippit"),
                        new XElement(
                            Drawing.dk1,
                            new XElement(
                                Drawing.sysClr,
                                new XAttribute(NoNamespace.val, "windowText"),
                                new XAttribute(NoNamespace.lastClr, "000000")
                            )
                        ),
                        new XElement(
                            Drawing.lt1,
                            new XElement(
                                Drawing.sysClr,
                                new XAttribute(NoNamespace.val, "window"),
                                new XAttribute(NoNamespace.lastClr, "FFFFFF")
                            )
                        ),
                        new XElement(
                            Drawing.dk2,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "1F1F1F"))
                        ),
                        new XElement(
                            Drawing.lt2,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "FFFFFF"))
                        ),
                        new XElement(
                            Drawing.accent1,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "4285F4"))
                        ),
                        new XElement(
                            Drawing.accent2,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "34A853"))
                        ),
                        new XElement(
                            Drawing.accent3,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "FBBC04"))
                        ),
                        new XElement(
                            Drawing.accent4,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "EA4335"))
                        ),
                        new XElement(
                            Drawing.accent5,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "A142F4"))
                        ),
                        new XElement(
                            Drawing.accent6,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "00ACC1"))
                        ),
                        new XElement(
                            Drawing.hlink,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "0563C1"))
                        ),
                        new XElement(
                            Drawing.folHlink,
                            new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "954F72"))
                        )
                    ),
                    new XElement(
                        Drawing.fontScheme,
                        new XAttribute(NoNamespace.name, "Clippit"),
                        new XElement(
                            Drawing.majorFont,
                            new XElement(Drawing.latin, new XAttribute(NoNamespace.typeface, "Aptos Display")),
                            new XElement(Drawing.ea, new XAttribute(NoNamespace.typeface, "")),
                            new XElement(Drawing.cs, new XAttribute(NoNamespace.typeface, ""))
                        ),
                        new XElement(
                            Drawing.minorFont,
                            new XElement(Drawing.latin, new XAttribute(NoNamespace.typeface, "Aptos")),
                            new XElement(Drawing.ea, new XAttribute(NoNamespace.typeface, "")),
                            new XElement(Drawing.cs, new XAttribute(NoNamespace.typeface, ""))
                        )
                    ),
                    new XElement(
                        Drawing.fmtScheme,
                        new XElement(
                            Drawing.fillStyleLst,
                            Enumerable.Repeat(
                                new XElement(
                                    Drawing.solidFill,
                                    new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "FFFFFF"))
                                ),
                                3
                            )
                        ),
                        new XElement(
                            Drawing.lnStyleLst,
                            Enumerable.Repeat(
                                new XElement(
                                    Drawing.ln,
                                    new XAttribute(NoNamespace.w, 9525),
                                    new XElement(
                                        Drawing.solidFill,
                                        new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "000000"))
                                    ),
                                    new XElement(Drawing.prstDash, new XAttribute(NoNamespace.val, "solid"))
                                ),
                                3
                            )
                        ),
                        new XElement(
                            Drawing.effectStyleLst,
                            Enumerable.Repeat(new XElement(Drawing.effectStyle, new XElement(Drawing.effectLst)), 3)
                        ),
                        new XElement(
                            Drawing.bgFillStyleLst,
                            Enumerable.Repeat(
                                new XElement(
                                    Drawing.solidFill,
                                    new XElement(Drawing.srgbClr, new XAttribute(NoNamespace.val, "FFFFFF"))
                                ),
                                3
                            )
                        )
                    )
                )
            )
        );
}
