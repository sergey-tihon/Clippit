# Clippy Memory

## Last Run
2026-06-26 16:28 UTC — Run 28251071047

## Comments Made
- #54: CRC32 improvement idea — explained reflection is fragile; suggested profiling first
- #67: Explained DocumentAssembler uses XPath 1.0; pointed to conditional row workaround
- #77: Explained limitations of DocumentAssembler re: figure captions/TOC update
- #103: Explained DocumentAssembler uses content controls not bookmarks
- #28: Design sketch for ExcelAssembler + follow-up linking to PR #165
- #66: Explained RegisterCustomHandler design + linked PR #163

## Open Clippy PRs
- #163: feat: RegisterCustomHandler extensibility API — open since March 2026
- #165: feat(excel): ExcelAssembler — open since March 2026
- #353: refactor: null patterns in Word/ files (8 files, 445 reps) — open 2026-06-26
- TBD (2026-06-26): test(common): FieldRetriever.ParseField unit tests FR001-FR010 (19 cases)
- TBD (2026-06-26): refactor(html): Html/ null patterns — HtmlToWmlCssApplier (63), HtmlToWmlCssParser (25), HtmlToWmlConverter (5) = 93 reps

## Backlog Cursor
All 6 non-activity open issues (#28, #54, #66, #67, #77, #103) have Clippy comments. No new human activity as of 2026-06-26.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (PR #346 merged 2026-06-23)
- Monthly Activity issue: June 2026 is #311
- PR #352 (null patterns in PtUtil+OxPtHelpers+WmlComparer files, 128 reps) MERGED 2026-06-26
- Null modernization for remaining files: OpenXmlRegex.cs (30), TextReplacer.cs (13), Word/Assembler/HtmlConverter.cs (12), small files

## Null Modernization Progress
Completed:
  - Excel converters #344 (merged)
  - FormattingAssembler.cs #345 (merged)
  - DocumentBuilder.cs #349 (merged)
  - ListItemRetriever.cs+HtmlToWmlConverterCore.cs+WmlComparer.ProduceDocument.cs #350 (merged)
  - PtUtil.cs+OxPtHelpers.cs+MetricsGetter.cs+PtOpenXmlDocument.cs+UnicodeMapper.cs+all WmlComparer files #352 (merged)
  - Word/ files (8 files, 445 reps) #353 open
  - Html/ files (3 files, 93 reps) TBD open 2026-06-26

Remaining files with == null / != null:
  - OpenXmlRegex.cs (30)
  - Internal/TextReplacer.cs (13)
  - Word/Assembler/HtmlConverter.cs (12)
  - Excel/PegBase.cs (26) — generated parser, may be auto-generated
  - Excel/SSFormula.cs (4)
  - Excel/XlsxTables.cs (14)
  - Comparer/CorrelatedSequence.cs (2)
  - small others (1 each): FluentPresentationBuilder.Copy.cs, MetricsGetter.cs, SmlCellFormatter.cs, StronglyTypedBlock.cs, PowerToolsBlockExtensions.cs, FieldRetriever.cs
  - HtmlToWmlConverterCore.cs (1 in JS comment — skip)
