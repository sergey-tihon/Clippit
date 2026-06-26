# Clippy Memory

## Last Run
2026-06-26 13:48 UTC — Run 28242095104

## Comments Made
- #54: CRC32 improvement idea — explained using reflection is fragile; suggested profiling first
- #67: Explained DocumentAssembler uses XPath 1.0; pointed to conditional row workaround
- #77: Explained limitations of DocumentAssembler re: figure captions/TOC update
- #103: Explained DocumentAssembler uses content controls not bookmarks
- #28: Design sketch for ExcelAssembler + follow-up linking to PR #165
- #66: Explained RegisterCustomHandler design + linked PR #163

## Open Clippy PRs
- #163: feat: RegisterCustomHandler extensibility API — open since March 2026
- #165: feat(excel): ExcelAssembler — open since March 2026
- TBD (created 2026-06-26): refactor: null patterns in Word/ files (ChartUpdater, RevisionProcessor, DocumentAssembler, MarkupSimplifier, ReferenceAdder, WmlDocument, WmlToHtmlConverter, WmlToXml) — 445 replacements across 8 files

## Backlog Cursor
All 6 non-activity open issues (#28, #54, #66, #67, #77, #103) have Clippy comments. No new human activity as of 2026-06-26.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (PR #346 merged 2026-06-23)
- Monthly Activity issue: June 2026 is #311
- PR #352 (null patterns in PtUtil+OxPtHelpers+WmlComparer files, 128 reps) MERGED 2026-06-26
- Null modernization (== null / != null series) COMPLETE for all files — Word/ files done in TBD PR 2026-06-26

## Null Modernization Progress
Completed (as/null → is not): RevisionProcessor#307, MarkupSimplifier#308, OpenXmlRegex#315, WmlToHtmlConverter+WmlToXml#336
Completed (== null / != null → is/is not):
  - FormattingAssembler.cs #345 (merged)
  - Excel converters #344 (merged)
  - DocumentBuilder.cs #349 (merged)
  - ListItemRetriever.cs+HtmlToWmlConverterCore.cs+WmlComparer.Private.Methods.ProduceDocument.cs #350 (merged)
  - PtUtil.cs+OxPtHelpers.cs+MetricsGetter.cs+PtOpenXmlDocument.cs+UnicodeMapper.cs+all WmlComparer files #352 (merged)
  - Word/ files (ChartUpdater+RevisionProcessor+DocumentAssembler+MarkupSimplifier+ReferenceAdder+WmlDocument+WmlToHtmlConverter+WmlToXml) TBD (445 reps, 2026-06-26)
Status: SERIES COMPLETE — all library files modernized
