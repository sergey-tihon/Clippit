# Clippy Memory

## Last Run
2026-05-31 16:00 UTC — Run 26717375823

## Comments Made
- #54: CRC32 improvement idea — explained using reflection is fragile; suggested profiling first
- #67: Explained DocumentAssembler uses XPath 1.0; pointed to conditional row workaround
- #77: Explained limitations of DocumentAssembler re: figure captions/TOC update
- #103: Explained DocumentAssembler uses content controls not bookmarks

## Open Clippy PRs
- #307: refactor: modernize `as` + null-check in RevisionProcessor (19 methods)
- #308: refactor: modernize `as` + null-check in MarkupSimplifier (8 methods + 1 lambda) — created this run
- #165: feat(excel): ExcelAssembler — open since March 2026
- #163: feat: RegisterCustomHandler extensibility API — open since March 2026

## Backlog Cursor
All 6 non-activity open issues have Clippy comments. No new human activity as of this run.
Issues: #28, #54, #66, #67, #77, #103

## Notes
- #273 (MetricsGetter modernize) merged
- #274 (TryGetValue perf) merged
- #275 (Word/Assembler helper files refactor) merged
- Monthly Activity issue: #249
- PR #307 CI: all checks passing (build ubuntu, build windows, generate-docs ✅)
- Next: Consider OpenXmlRegex.cs, WmlToXml.cs, WmlToHtmlConverter.cs for similar modernization
