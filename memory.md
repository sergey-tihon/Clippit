# Clippy Memory

## Last Run
2026-06-26 04:23 UTC — Run 28216934908

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
- TBD (created 2026-06-26): refactor: null patterns in PtUtil.cs, OxPtHelpers.cs, PtOpenXmlUtil.cs, MetricsGetter.cs, PtOpenXmlDocument.cs, UnicodeMapper.cs (75 replacements)
- TBD (created 2026-06-26): chore(deps): TUnit 1.56.25 → 1.56.35

## Backlog Cursor
All 6 non-activity open issues (#28, #54, #66, #67, #77, #103) have Clippy comments. No new human activity as of 2026-06-26.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (PR #346 merged 2026-06-23)
- Monthly Activity issue: June 2026 is #311
- PR #350 (null patterns remaining 3 files, 2026-06-25) MERGED
- Null modernization REMAINING files: WmlComparer files (~70 patterns across 7 files)
  - WmlComparer.Internal.Methods.ComparisonUnits.cs (~7)
  - WmlComparer.Private.Methods.Lcs.cs (~12)
  - WmlComparer.Private.Methods.Hashing.cs (~6)
  - WmlComparer.Public.Methods.Consolidate.cs (~28)
  - WmlComparer.Public.Methods.Compare.cs (~3)
  - WmlComparer.Public.Methods.GetRevisions.cs (~1)
  - WmlComparer.Private.Methods.PreProcessMarkup.cs (~13)
  - FieldRetriever.cs, Word/DocumentAssembler.cs, Word/RevisionProcessor.cs, Word/MarkupSimplifier.cs (have some patterns remaining too)

## Null Modernization Progress
Completed (as/null → is not): RevisionProcessor#307, MarkupSimplifier#308, OpenXmlRegex#315, WmlToHtmlConverter+WmlToXml#336
Completed (== null / != null → is/is not):
  - FormattingAssembler.cs #345 (merged)
  - Excel converters #344 (merged)
  - DocumentBuilder.cs #349 (merged)
  - ListItemRetriever.cs+HtmlToWmlConverterCore.cs+WmlComparer.Private.Methods.ProduceDocument.cs #350 (merged)
  - PtUtil.cs+OxPtHelpers.cs+MetricsGetter.cs+PtOpenXmlDocument.cs+UnicodeMapper.cs+PtOpenXmlUtil.cs (TBD 2026-06-26)
Remaining: WmlComparer files (~70 patterns), FieldRetriever.cs, plus scattered patterns in Word/ files
