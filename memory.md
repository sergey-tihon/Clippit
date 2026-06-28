# Clippy Memory

## Last Run
2026-06-28 16:03 UTC — Run 28327906587

## Comments Made
- #54: CRC32 improvement idea — explained reflection is fragile; suggested profiling first
- #67: Explained DocumentAssembler uses XPath 1.0; pointed to conditional row workaround
- #77: Explained limitations of DocumentAssembler re: figure captions/TOC update
- #103: Explained DocumentAssembler uses content controls not bookmarks
- #28: Design sketch for ExcelAssembler + follow-up linking to PR #165
- #66: Explained RegisterCustomHandler design + linked PR #163

## Open Clippy PRs
- #163: feat: RegisterCustomHandler extensibility API — rebased on master 2026-06-28, CI triggered
- #165: feat(excel): ExcelAssembler — rebased on master 2026-06-28, CI triggered
- TBD: chore(deps): update TUnit 1.56.35 → 1.57.0 — created 2026-06-28

## Backlog Cursor
All 6 non-activity open issues (#28, #54, #66, #67, #77, #103) have Clippy comments. No new human activity as of 2026-06-28.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (PR #346 merged 2026-06-23)
- Monthly Activity issue: June 2026 is #311
- Null modernization series COMPLETE for all core library files
  - Only Excel/PegBase.cs (external library, CPOL license) was intentionally excluded
- PR #354 (FieldRetriever tests), #355 (Html/ null), #356 (batch5) all MERGED 2026-06-28

## Null Modernization Progress - COMPLETE
All batches merged:
  - #344 Excel converters, #345 FormattingAssembler.cs, #349 DocumentBuilder.cs
  - #350 ListItemRetriever+HtmlToWmlConverterCore+WmlComparer.ProduceDocument
  - #352 PtUtil+OxPtHelpers+MetricsGetter+PtOpenXmlDocument+UnicodeMapper+all WmlComparer files
  - #353 Word/ root files (8 files, 445 reps)
  - #355 Html/ files (3 files, 93 reps)
  - #356 final batch (10 files: OpenXmlRegex, TextReplacer, HtmlConverter, PtOpenXmlUtil, XlsxTables,
         SSFormula, SmlCellFormatter, PowerToolsBlockExtensions, StronglyTypedBlock,
         FluentPresentationBuilder.Copy) + Lazy<T> KnownFamilies + ??= improvements

## Future Ideas
- ExcelAssembler row-repetition support (follow-up to #165)
- Thread-safety in static caches (UnknownFonts HashSet — noted in batch5, may need follow-up)
