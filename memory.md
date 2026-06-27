# Clippy Memory

## Last Run
2026-06-27 16:02 UTC — Run 28294305804

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
- #354: test(common): FieldRetriever.ParseField unit tests FR001-FR010 — open 2026-06-26
- #355: refactor(html): Html/ null patterns — open 2026-06-26
- TBD (2026-06-27): refactor batch5: 10 files (OpenXmlRegex, TextReplacer, HtmlConverter, PtOpenXmlUtil, XlsxTables, SSFormula, SmlCellFormatter, PowerToolsBlockExtensions, StronglyTypedBlock, FluentPresentationBuilder.Copy) + Lazy<T> KnownFamilies fix

## Backlog Cursor
All 6 non-activity open issues (#28, #54, #66, #67, #77, #103) have Clippy comments. No new human activity as of 2026-06-27.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (PR #346 merged 2026-06-23)
- Monthly Activity issue: June 2026 is #311
- Null modernization series COMPLETE for all core library files
  - Only Excel/PegBase.cs (external library, CPOL license) was intentionally excluded
- PR #353 (Word/ null patterns, 445 reps) MERGED 2026-06-26

## Null Modernization Progress - COMPLETE
All batches:
  - #344 Excel converters (merged)
  - #345 FormattingAssembler.cs (merged)
  - #349 DocumentBuilder.cs (merged)
  - #350 ListItemRetriever+HtmlToWmlConverterCore+WmlComparer.ProduceDocument (merged)
  - #352 PtUtil+OxPtHelpers+MetricsGetter+PtOpenXmlDocument+UnicodeMapper+all WmlComparer files (merged)
  - #353 Word/ root files (8 files, 445 reps) (merged)
  - #355 Html/ files (3 files, 93 reps) (open)
  - Batch5: 10 files (OpenXmlRegex, TextReplacer, HtmlConverter, PtOpenXmlUtil, XlsxTables, SSFormula, SmlCellFormatter, + small Core/PowerPoint files) + Lazy<T> KnownFamilies fix (open, created 2026-06-27)

## Future Ideas
- ExcelAssembler row-repetition support (follow-up to #165)
- Look at thread-safety in other static caches (UnknownFonts HashSet is still not thread-safe)
