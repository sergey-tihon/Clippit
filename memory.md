# Clippy Memory

## Last Run
2026-06-25 16:31 UTC — Run 28185041665

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
- TBD (created 2026-06-25): refactor: null patterns in ListItemRetriever.cs, HtmlToWmlConverterCore.cs, WmlComparer.Private.Methods.ProduceDocument.cs (367 replacements)

## Backlog Cursor
All 6 non-activity open issues (#28, #54, #66, #67, #77, #103) have Clippy comments. No new human activity as of 2026-06-25.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (PR #346 merged 2026-06-23)
- Monthly Activity issue: June 2026 is #311
- Null modernization series COMPLETE: no more == null / != null patterns in main target files
  - Done: ExcelConverters (#344), FormattingAssembler (#345), DocumentBuilder (#349), ListItemRetriever+HtmlToWmlConverterCore+WmlComparer.ProduceDocument (TBD 2026-06-25)
  - Other files (OxPtHelpers.cs, PtOpenXmlDocument.cs, etc.) still have some patterns but are lower priority

## Null Modernization Progress
Completed (as/null → is not): RevisionProcessor#307, MarkupSimplifier#308, OpenXmlRegex#315, WmlToHtmlConverter+WmlToXml#336
Completed (== null / != null → is/is not): FormattingAssembler.cs #345, Excel converters #344, DocumentBuilder.cs #349, ListItemRetriever.cs+HtmlToWmlConverterCore.cs+WmlComparer.Private.Methods.ProduceDocument.cs (TBD 2026-06-25)
Remaining: OxPtHelpers.cs (~7), PtOpenXmlDocument.cs (~7), MetricsGetter.cs (~15), PtOpenXmlUtil.cs (~5) — lower priority scattered patterns
