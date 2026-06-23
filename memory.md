# Clippy Memory

## Last Run
2026-06-23 01:03 UTC — Run 27994786796

## Comments Made
- #54: CRC32 improvement idea — explained using reflection is fragile; suggested profiling first
- #67: Explained DocumentAssembler uses XPath 1.0; pointed to conditional row workaround
- #77: Explained limitations of DocumentAssembler re: figure captions/TOC update
- #103: Explained DocumentAssembler uses content controls not bookmarks

## Open Clippy PRs
- #163: feat: RegisterCustomHandler extensibility API — open since March 2026
- #165: feat(excel): ExcelAssembler — open since March 2026
- #342: test(cli): fix duplicate CLI test codes (ExcelVerify CLI051-055→CLI078-082, ExcelToHtml CLI061-066→CLI083-088) — created 2026-06-22
- #343: chore(deps): TUnit 1.54.0→1.56.25 + JsonSchema.Net 9.2.1→9.2.2 — created 2026-06-22
- TBD#: refactor: FormattingAssembler.cs null-pattern modernization (252 replacements) — created 2026-06-23
- TBD#: refactor: Excel converters null-pattern modernization (227 replacements) — created 2026-06-23

## Backlog Cursor
All 6 non-activity open issues (#28, #54, #66, #67, #77, #103) have Clippy comments. No new human activity as of 2026-06-23.

## Notes
- v3.5.0 released 2026-06-20: SkiaSharp replaces SixLabors.ImageSharp.Drawing (PR #341)
- Monthly Activity issue: June 2026 is #311
- PR #333 (NikiforovAll Strict OOXML) merged 2026-06-16
- PR #336 (WmlToHtmlConverter is-not) merged 2026-06-13
- PR #337 (TUnit 1.54.0) merged 2026-06-13
- System.CommandLine pinned at 3.0.0-preview.4 (intentional - pre-release)
- Null-pattern modernization remaining: DocumentBuilder.cs (189), ListItemRetriever.cs (95), HtmlToWmlConverterCore.cs (163), WmlComparer.Private.Methods.ProduceDocument.cs (89)

## Null Modernization Progress
Completed (as/null → is not): RevisionProcessor#307, MarkupSimplifier#308, OpenXmlRegex#315, WmlToHtmlConverter+WmlToXml#336
Completed (== null / != null): FormattingAssembler.cs (TBD#, 2026-06-23), Excel converters (TBD#, 2026-06-23)
