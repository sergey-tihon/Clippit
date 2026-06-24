# Clippy Memory

## Last Run
2026-06-24 16:29 UTC — Run 28113570461

## Comments Made
- #54: CRC32 improvement idea — explained using reflection is fragile; suggested profiling first
- #67: Explained DocumentAssembler uses XPath 1.0; pointed to conditional row workaround
- #77: Explained limitations of DocumentAssembler re: figure captions/TOC update
- #103: Explained DocumentAssembler uses content controls not bookmarks
- #28: Design sketch for ExcelAssembler + follow-up linking to PR #165

## Open Clippy PRs
- #163: feat: RegisterCustomHandler extensibility API — open since March 2026
- #165: feat(excel): ExcelAssembler — open since March 2026
- TBD#(#aw_db_null): refactor: DocumentBuilder.cs null-pattern modernization (189 replacements) — created 2026-06-24

## Backlog Cursor
All 6 non-activity open issues (#28, #54, #66, #67, #77, #103) have Clippy comments. No new human activity as of 2026-06-24.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (PR #346 merged 2026-06-23)
- PR #347: bump actions/checkout 6→7 (merged 2026-06-23)
- Monthly Activity issue: June 2026 is #311
- System.CommandLine now at preview.5 (was intentionally pinned at preview.4 per previous note)

## Null Modernization Progress
Completed (as/null → is not): RevisionProcessor#307, MarkupSimplifier#308, OpenXmlRegex#315, WmlToHtmlConverter+WmlToXml#336
Completed (== null / != null → is/is not): FormattingAssembler.cs #345, Excel converters #344, DocumentBuilder.cs (TBD/#aw_db_null, 2026-06-24)
Remaining: ListItemRetriever.cs (95), HtmlToWmlConverterCore.cs (163), WmlComparer.Private.Methods.ProduceDocument.cs (89)
