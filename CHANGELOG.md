# Changelog

## [3.6.0] - July 3, 2026

- feat(cli): add `word compare` command for DOCX diff (#365)
- feat(cli): add Windows ARM64 npm binary (#363)
- feat(cli): add Linux ARM64 npm binary (#360)
- perf: use `FrozenSet<T>` for `HashSet` lookups (#361)
- perf: convert `Dictionary` to `FrozenDictionary` (#362)
- perf: replace remaining `List<T>` with `FrozenSet<T>` in `PtOpenXmlUtil` (#371)
- fix: guard `TextReplacer` against `IndexOutOfRangeException` on empty replacement string (#381)
- fix: guard `WmlComparer.HashBlockLevelContent` against `ArgumentNullException` from missing `pt:Unid` attributes (#380)
- fix: guard `WmlComparer.FindIndexOfNextParaMark` against `InvalidCastException` from non-`ComparisonUnitWord` units (#385)
- fix(html): preserve `w:anchor` fragment on external hyperlinks — append `#anchor` when the URL has no existing fragment (#384)
- fix(html): correct `rowspan` for vertically merged cells when preceding cells use `w:gridSpan` (#383)
- fix(pptx): handle `ExtendedPart` chart workbooks in `CopyChartObjects` and `CopyExtendedChartObjects` — prevents silent data loss when a chart's external data relationship resolves to an `ExtendedPart` (#388)
- refactor: modernize `== null`/`!= null` to `is null`/`is not null` across 20+ files (#349, #350, #352, #353, #355, #356)
- test: add `FieldRetriever.ParseField` unit tests FR001–FR010 (#354)
- test: add regression tests for `TextReplacer` empty-string replacement TR007–TR008 (#381)
- test: add regression tests for `WmlComparer` defensive fixes WC380, WC385 (#380, #385)
- chore(deps): update TUnit 1.56.35 → 1.57.17 (#351, #359, #369)
- chore(ci): update `actions/checkout` v6 → v7 (#347)

## [3.5.1] - June 23, 2026

- chore(deps): update SkiaSharp 3.119.4 → 4.148.0 with migration fix (#346)
- chore(deps): update Microsoft.NET.Test.Sdk 18.6.0 → 18.7.0 (#346)
- chore(deps): update dotnet-outdated-tool 4.7.1 → 4.8.1 (#346)
- chore(deps): update csharpier 1.2.6 → 1.3.0 (#346)
- chore(deps): update System.CommandLine 3.0.0-preview.4 → preview.5 (#346)
- fix: replace deprecated `SKPaint.MeasureText` with `SKFont.MeasureText` for SkiaSharp 4.x compatibility (#346)

## [3.5.0] - June 20, 2026

- feat: replace SixLabors.ImageSharp with SkiaSharp for all image processing (#341)
- feat: add Linux native SkiaSharp asset support (`SkiaSharp.NativeAssets.Linux.NoDependencies`) (#341)
- fix(html): guard null `srcAttribute` in `LoadImageForTransform` to prevent `NullReferenceException` (#341)
- fix(html): encode image before adding `ImagePart` to prevent dangling empty relationship on encode failure (#341)
- fix(docs): update `WmlToHtmlConverter` tutorial from ImageSharp to SkiaSharp API (#341)
- fix(docs): add `using SkiaSharp;` note in tutorial code sample (#341)
- refactor: replace `Image.Load`/`Image.Save` with `SKBitmap.Decode`/`SKImage.Encode` across all converters and helpers (#341)
- refactor: replace `SixLabors.Fonts` text measurement with `SKTypeface`/`SKPaint.MeasureText` in `MetricsGetter` and `PtOpenXmlUtil` (#341)
- refactor: use safe fallback chain for `GetTextWidth` — `ArgumentException` catch now works correctly after removing blanket catch in `_getTextWidth` (#341)
- fix(cli): use `FileMode.Create` instead of `FileMode.OpenOrCreate` in `DefaultImageHandler` to avoid corrupted image on re-encode (#341)
- fix(cli): return `null` instead of `null!` in `CreateInlineImage` failure paths (#341)
- fix(test): add explicit null-check with `InvalidOperationException` in `BuildTestPng` SkiaSharp encode (#341)
- perf: wrap `SKTypeface`/`SKFont`/`SKImage` with `using` to prevent native handle leaks (#341)
- chore(deps): remove `SixLabors.ImageSharp.Drawing` dependency (#341)
- chore(deps): add `SkiaSharp 3.119.4` (#341)

## [3.4.7] - June 16, 2026

- feat(cli): add `word to-html`, `word from-html`, and `excel to-html` commands (#338)
- fix(powerpoint): support ISO/IEC 29500 Strict OOXML presentations — translate Strict namespace URIs lazily via `XmlReader` wrapper in `GetXDocument()` (#339)
- fix(powerpoint): translate `<a:graphicData uri=…>` content-type URIs and fix source `XDocument` mutation in `CopyPresentationParts` (#340)
- refactor: modernize `as` + null-check to `is not` pattern matching in `WmlToHtmlConverter` and `WmlToXml` (#336)
- chore(deps): update TUnit 1.49.0 → 1.54.0 (#337)

## [3.4.6] - June 10, 2026

- feat(cli): add `word verify` and `excel verify` commands (#329)
- fix(html): convert `w:pageBreakBefore` paragraph property to CSS `page-break-before` (#325)
- refactor: modernize `as` + null-check to `is not` pattern matching in `OpenXmlRegex` (#315)
- chore(deps): update TUnit 1.48.6 → 1.49.0 (#327)
- chore(deps): update JsonSchema.Net 7.4.0 → 9.2.1 (#317)

## [3.4.5] - June 3, 2026

- feat: add Clippit CLI — `pptx split`, `pptx build`, `pptx verify`, and `version` commands (#309)
- fix(pptx): fix PowerPoint relationship copying edge cases — prevent missing parts after slide copy (#312)
- fix(cli): set `AssemblyName=clippit` to avoid NativeAOT rename on Windows
- fix(cli): use `$(ExeSuffix)` in `RenameNativeAotExecutable` MSBuild target
- fix(cli): fix npm publishing workflow
- refactor: modernize `as` + null-check to `is not` pattern matching in `RevisionProcessor` (#307)
- refactor: modernize `as` + null-check to `is not` pattern matching in `MarkupSimplifier` (#308)
- test(excel): expand `WorksheetAccessor` tests — `SetCellValue`/`GetCellValue` round-trips WA004–WA013 (#314)
- chore(deps): update TUnit 1.47.0 → 1.48.6 (#313)
- docs: add README for Clippit.Cli NuGet tool and npm clippit package (#320)
- docs: restore classic Clippy hero image and redesign homepage with retro Clippy speech bubble

## [3.4.4] - May 20, 2026

- fix(pptx): prevent `NullReferenceException` in `GetSlideTitle` when `cSld`/`spTree` is missing (#287)
- fix(html): render Word page breaks as CSS `page-break` div in HTML output (#281)
- fix(word): `RemoveSoftHyphens` and `RemoveLastRenderedPageBreak` in `MarkupSimplifier` were silently ignored (#284)
- perf: replace `ToUpper()`+`StartsWith()` with `OrdinalIgnoreCase` comparisons (#282)
- perf: replace `.Where(pred).Count()` with `.Count(pred)` and `Count()!=1` with `Skip(1).Any()` (#270)
- perf: replace `ContainsKey`+indexer double-lookups with `TryGetValue` (#274)
- refactor: modernize `GetListItemText` locale files (#283)
- refactor: modernize `UnicodeMapper.cs` (#278)
- refactor: modernize Word/Assembler helper files (#275)
- refactor: modernize `MetricsGetter.cs` — file-scoped namespace, `is null`/`is not null`, pattern matching (#273)
- refactor: modernize `FieldRetriever.cs` — file-scoped namespace, `W.*` constants, `is null`, slices, `TryGetValue` (#272)
- refactor: simplify `Select().SelectMany()` chains in `WmlComparer` to `SelectMany()` (#271)
- refactor: simplify `FileUtils.GetFilesRecursive` using `SearchOption.AllDirectories` (#268)
- refactor: modernize null checks to use `ArgumentNullException.ThrowIfNull` (#265)
- test: add unit tests for `OpenXmlPowerToolsDocument` and `OpenXmlMemoryStreamDocument` (OXD001–OXD040) (#266)
- test: add `ReferenceAdder` unit tests (RA100–RA140) and fix `AddTof` `w:dirty` namespace bug (#263)
- chore(deps): update TUnit 1.43.11 → 1.45.8 (#264, #267, #276, #277)

## [3.4.3] - May 6, 2026

- fix(pptx): fix `ArgumentNullException` in `FluentPresentationBuilder` when slides have null data (#256)
- fix(word): rely on SDK 3.x built-in handling for invalid hyperlink URIs — removes manual `UriFixer` patching (#253)
- perf: replace LINQ char filter with `ReplaceLineEndings` in `FlatOpc.ToPackage` (#250)
- chore(deps): update TUnit 1.41.0 → 1.43.11 (#254)

## [3.4.2] - April 30, 2026

- fix(pptx): make all part copies resilient to corrupt ZIP local file headers (#234)
- refactor: replace ContainsKey+indexer double-lookups with TryGetValue (#246)
- refactor(excel): remove dead code from SpreadsheetDocumentManager + add WorksheetAccessor unit tests (#247)
- refactor: modernize UriFixer — convenience overload, leaveOpen param, eliminate test duplication (#237)
- refactor: modernise random/guid usage, fix discarded Guid bug, and remove dead code (#244)
- test(excel): add SmlDataRetriever unit tests (SDR001–SDR025) (#236)
- chore(deps): update TUnit 1.37.10 → 1.40.10 (#235, #242, #243, #245)

## [3.4.1] - April 23, 2026

- fix: resolve `ArgumentException` for slides with `custData`+`CustomXmlProperties` (#232)
- refactor: replace generic `Exception` with specific exception types (#228)
- perf: eliminate redundant allocations in `MakeValidXml` and `AddIfMissing` (#229)
- ci: add NuGet publish workflow triggered by version tags (#227)
- test: add unit tests for `Base64.ChunkBase64` and `ConvertFromBase64` (#230)
- chore(deps): update TUnit 1.37.0 → 1.37.10 (#222)

## [3.4.0] - April 19, 2026

- feat: add `Optional="true"` support to `<Table>` directive in DocumentAssembler (#150)
- feat: add `FitWithin` image sizing mode to DocumentAssembler (#168)
- feat: add `RelationshipValidator` to detect dangling `r:id` references in OpenXml parts (#160)
- feat: extract shared `RomanNumeralUtil` with `ToUpperRoman`/`ToLowerRoman` helpers (#164)
- feat(html): render floating tables (`w:tblpPr`) with CSS float in `WmlToHtmlConverter` (#180)
- fix(word): correct Russian list item text for values ≥ 100 (#187)
- fix(word): guard `cardinalText`/`ordinalText` against out-of-range `levelNumber` in multiple locales (#200, #201, #208)
- fix(html): render DOCX text boxes as inline-block divs in `WmlToHtmlConverter` (#166)
- perf(pptx): single-pass dispatch in `CopyRelatedPartsForContentParts` + lazy `SaveAndCleanup` (#175)
- perf(pptx): avoid `OuterXml`→`Parse` roundtrip in `SlidePartData.GetShapeDescriptor` (#178)
- perf: convert `Regex` instances to source-generated `[GeneratedRegex]` in several modules (#196)
- perf(comparer): reduce heap allocations in `WmlComparerUtil` SHA-1 hashing (#199)
- perf: modernize SHA1/SHA256 hashing — use `SHA1.HashData()`, `SHA256.HashData()`, `Convert.ToHexString()` (#184)
- perf: eliminate per-character heap allocations in base64 chunking (FlatOpc / Base64) (#209)
- refactor: remove custom relationship ID generator — delegate to OpenXML SDK APIs (#161)
- refactor: simplify LINQ patterns across Excel, Word, and Html modules (#207)
- docs: improve README with NuGet badges, project overview, quick-start, and docs link (#183)
- docs: switch DocFX documentation site to modern Bootstrap-based theme (#211)
- test(word): add unit tests for `ListItemTextGetter_Default`, fr_FR, sv_SE, zh_CN, tr_TR, ru_RU locales (#198, #206)
- test(word): add MS004–MS009 for `MarkupSimplifier` settings coverage (#213)
- test(excel): add `ParseFormula` and `XlsxTables` cell-address utility tests (#169, #185)
- test(common): add `TextReplacer` unit tests (#167)

## [3.3.1] - March 24, 2026

- fix: handle dangling r:id on p:oleObj/p:externalData — KeyNotFoundException in slide publishing (#156)
- fix: resolve Content directives in v:textpath/`@string` for VML watermarks (#141)

## [3.3.0] - March 22, 2026

- fix: prevent malformed XLSX when worksheet has no data rows (#108)
- fix: use min-width for tab-preceding span to prevent text overflow (#110)
- fix: update docProps/app.xml metadata in PublishSlides output (#114)
- fix: update numFmts count attribute when adding custom number formats (#124)
- fix: custom numFmt IDs must start at 164 per ECMA-376 (#131)
- fix: TextReplacer.CloneWithAnnotation now returns the clone, not the original (#135)
- perf: `O(1)` dictionary lookup for media deduplication cache in FluentPresentationBuilder (#128)
- perf: cache next slide ID in FluentPresentationBuilder to avoid `O(n²)` scan (#140)
- perf: cache compiled Regex, use `Any()` over `Count() == 0`, use `Element()` over `Elements().First()` (#143)
- refactor: modernise assembler helper classes (#118)
- refactor: modernize DocumentAssembler.cs with collection expressions and catch cleanup (#138)
- hk: DocumentFormat.OpenXml 3.5.1, TUnit 1.20.0 (#146)
- hk: Modernize CI workflow and update dependencies (#121)

## [3.2.0] - January 25, 2026

- hk: DocumentFormat.OpenXml 3.4.1

## [3.1.0] - August 30, 2025

- Build with .NET 10.0 and provide .NET 8 & 10 binaries (#106)

## [3.0.2] - August 30, 2025

- Extends `DocumentAssembler` to also apply HTML formatting on content in table cells (#101)

## [3.0.1] - July 30, 2025

- fix: incorrect insertion of sldLayoutIdLst element for SlideMaster (#100)

## [3.0.0] - June 30, 2025

- Public exposure of `IFluentPresentationBuilder` ergonomic API for the code behind `PresentationBuilder`

## [2.5.2] - June 28, 2025

- fix: relationship of theme parts (#99)

## [2.5.1] - March 30, 2025

- Issue 95 and other issues #96 by @MalcolmJohnston
- hk: DocumentFormat.OpenXml 3.3.0

## [2.5.0] - February 7, 2025

- Insert Documents or Document Templates using Document Assembler #93 by @MalcolmJohnston

## [2.4.2] - January 15, 2025

- fix: Font Sizing Font Styling and Font Coloring is not working #94
- hk: dependencies update

## [2.4.1] - January 4, 2025

- fix: Bug fix/document assembler paragraph properties not copied from template #91
- fix: slide naming issues during publishing presentation with \*.potx extension
- hk: Migration to xUnit.v3

## [2.4.0] - December 8, 2024

- feat: Added support for HTML fragments in Document Assembler Content tag (#86)

## [2.3.0] - November 23, 2024

- hk: build with .NET 9.0
- hk: DocumentFormat.OpenXml 3.2.0
- hk: Pack with `dotnet pack` instead of `paket`
- feat: handle invalid xml characters (#89)

## [2.2.2] - November 1, 2024

- feat: DocumentFormat.OpenXml 3.1.0 -> 3.1.1
- docs: Fix document assembler images and other tidy up (#84)
- hk: Document Assembler - Small Code Reorganisation (#83)

## [2.2.1] - October 8, 2024

- fix: optimized memory usage during slide publishing

## [2.2.0] - September 8, 2024

- DocumentFormat.OpenXml 3.0.2 -> 3.1.0
- DocumentFormat.OpenXml.Framework 3.0.2 -> 3.1.0
- SixLabors.ImageSharp 3.1.4 -> 3.1.5
- SixLabors.ImageSharp.Drawing 2.1.3 -> 2.1.4

## [2.1.1] - April 18, 2024

- fix: regression cased by migration to struct inside OpenXML v3 #75
- Dependencies updated

## [2.1.0] - Mar 18, 2024

- Migration to .NET 8.0

## [2.0.1] - Jan 16, 2024

- Fix regression cased by migration to struct inside OpenXML v3
- DocumentFormat.OpenXml 3.0 -> 3.0.1
- SixLabors.Fonts 2.0 -> 2.0.1
- SixLabors.ImageSharp 3.0.2 -> 3.1.2
- SixLabors.ImageSharp.Drawing 2.0.1 -> 2.1.0

## [2.0.0] - Nov 26, 2023

- Migration to .NET 6.0
- DocumentFormat.OpenXml 3.0.0-beta0003
- `System.Drawing.Common` replaced by SixLabors.ImageSharp.Drawing`
- Removed `libgdiplus` dependency
- Drop old `WmlComparer`

## [1.13.5] - Mar 27, 2023

- perf: improved PublishSlides perf #61

## [1.13.4] - Dec 6, 2022

- fix: issues with 7.0 packages

## [1.13.3] - Dec 5, 2022

- fix: PublishSlides: missing ppt/metadata for google presentations #59
- Dependencies updated

## [1.13.2] - Oct 4, 2022

- remove caracters to prevent > at document assembler [#57](https://github.com/sergey-tihon/Clippit/pull/57)
- Handle W.lastRenderedPageBreak in UnicodeMapper [#58](https://github.com/sergey-tihon/Clippit/pull/58)

## [1.13.1] - Sept 27, 2022

- `long` overload for `Cell.Number`

## [1.13.0] - Sept 27, 2022

- DocumentFormat.OpenXml (2.18)
- Added Cell.Bool for Excel helpers

## [1.12.2] - Jul 13, 2022

- fix: PtOpenXmlUtil: process corrupted OpenXmlPart items [#56](https://github.com/sergey-tihon/Clippit/pull/56)

## [1.12.1] - Jul 11, 2022

- DocumentFormat.OpenXml (2.17.1)

## [1.12.0] - Jun 2, 2022

- Static Cell builder class for simpler and safer Excel generation
- Auto new modified date for decks composed with PresentationBuilder

## [1.11.0] - May 29, 2022

- PublishSlides: reduced memory consumption [#53](https://github.com/sergey-tihon/Clippit/pull/53)

## [1.10.2] - Mar 16, 2022

- DocumentFormat.OpenXml (2.16.0)
- Lock `System.Drawing.Common` version to `v5`

## [1.10.0] - Feb 19, 2022

- Added support of ExtendedChartPart (<http://schemas.microsoft.com/office/drawing/2014/chartex>) in PresentationBuilder and DocumentBuilder.

## [1.9.2] - Feb 17, 2022

- Fix: Don't show hidden slides after PresentationBuilder.BuildPresentation

## [1.9.1] - Feb 9, 2022

- Fixed incorrect usage of Stream API

## [1.9.0] - Jan 28, 2022

- DocumentFormat.OpenXml (2.15.0)

## [1.8.2] - Dec 16, 2021

- Improved memory consumption for PresentationBuilder.PublishSlides [#44 by @f1nzer](https://github.com/sergey-tihon/Clippit/pull/44)
- Revert DocumentFormat.OpenXml back to 2.13.1 (because of [this issue](https://github.com/OfficeDev/Open-XML-SDK/issues/1069))

## [1.8.1] - Nov 5, 2021

- DocumentFormat.OpenXml (2.14.0)

## [1.8.0] - Oct 25, 2021

- DocumentAssembler: Support for multi-value XPath results [#39](https://github.com/sergey-tihon/Clippit/pull/39)

## [1.7.3] - Oct 7, 2021

- Fixed copy Chart Style Parts in FluentPresentationBuilder

## [1.7.2] - Aug 23, 2021

- DocumentFormat.OpenXml v2.13.1

## [1.7.1] - Aug 18, 2021

- Resolving bug with nested rowspans [#35](https://github.com/sergey-tihon/Clippit/pull/35)

## [1.7.0] - Aug 15, 2021

- DocumentAssembler: Support for images [#31](https://github.com/sergey-tihon/Clippit/pull/31)

## [1.6.1] - Jul 11, 2021

- New docs site generated by DocFx [#33](https://github.com/sergey-tihon/Clippit/pull/33)
- Tests suite cleanup, samples converted to tests [#32](https://github.com/sergey-tihon/Clippit/pull/32)

## [1.6.0] - Jul 3, 2021

- Add WorkbookDfn.WriteTo(Stream) method [#29](https://github.com/sergey-tihon/Clippit/pull/29)
- Fixed generation of multiple Excel tables
- Fixed formatting of cells with DateTime

## [1.5.0] - Jun 26, 2021

- Structural comparison for Theme, Master, Layout [#20](https://github.com/sergey-tihon/Clippit/pull/20)
- Auto-scaling for slides from presentations with different slide size
- DocumentFormat.OpenXml (2.13)
- System.Drawing.Common (5.0.2)

## [1.4.0] - Feb 10, 2021

- DocumentBuilder: Added ISource and TableCellSource - [#17](https://github.com/sergey-tihon/Clippit/pull/17)
- DocumentFormat.OpenXml (2.12.1)
- Added dependency on System.IO.Packaging
- Fixed DateTime/DateTimeOffset serialisation format to Excel

## [1.3.1] - July 30, 2020

- PresentationBuilder: Fixed CopyExtendedPart

## [1.3.0] - July 29, 2020

- DocumentFormat.OpenXml (2.11.3)
- PresentationBuilder: Bug Fixes [#16](https://github.com/sergey-tihon/Clippit/pull/16)

## [1.2.1] - May 2, 2020

- HTML to WML: Allow font-size with unit rem [#13](https://github.com/sergey-tihon/Clippit/pull/13)

## [1.2.0] - March 12, 2020

- DocumentFormat.OpenXml (2.10.1)

## [1.1.4] - December 6, 2019

- Slide title extracted and saved to document title [#10](https://github.com/sergey-tihon/Clippit/pull/10)
- Ensured that Modified date is propagated
- Compiled and tested on .NET Core 3.1 [#9](https://github.com/sergey-tihon/Clippit/pull/9)

## [1.1.3] - December 3, 2019

- Don't rename theme when we extract slides from auto-generated 1-layout master

## [1.1.2] - December 2, 2019

- Bug fixes

## [1.1.1] - December 1, 2019

- Extract and merge presentation without slides (only masters) [#5](https://github.com/sergey-tihon/Clippit/pull/5)

## [1.1.0] - November 29, 2019

- Added Slide Publishing API [#2](https://github.com/sergey-tihon/Clippit/pull/2)

## [1.0.0] - November 21, 2019

- Initial release: [latest version of OpenXmlPowerTools](https://github.com/EricWhiteDev/Open-Xml-PowerTools/tree/6e56a5f5cf662f3bd3da87945a5d3ed2329964ff)
