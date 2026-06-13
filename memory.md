# Clippy Memory

## Last Run
2026-06-13 10:11 UTC — Run 27463338531

## Comments Made
- #54: CRC32 improvement idea — explained using reflection is fragile; suggested profiling first
- #67: Explained DocumentAssembler uses XPath 1.0; pointed to conditional row workaround
- #77: Explained limitations of DocumentAssembler re: figure captions/TOC update
- #103: Explained DocumentAssembler uses content controls not bookmarks
- #333: Reviewed NikiforovAll's PR: fix(PowerPoint) Strict OOXML — fix is correct, tests solid, minor efficiency observation

## Open Clippy PRs
- #163: feat: RegisterCustomHandler extensibility API — open since March 2026
- #165: feat(excel): ExcelAssembler — open since March 2026
- TBD#: refactor WmlToHtmlConverter/WmlToXml is-not patterns — created 2026-06-13
- TBD#: chore(deps) TUnit 1.49.0 → 1.54.0 — created 2026-06-13

## Backlog Cursor
All 6 non-activity open issues have Clippy comments. No new human activity as of this run.
Issues: #28, #54, #66, #67, #77, #103
PR #333 (NikiforovAll): reviewed 2026-06-13

## Notes
- v3.4.6 released 2026-06-10 by sergey-tihon
- SixLabors.ImageSharp.Drawing 3.0.0 requires paid commercial license - do NOT upgrade
- System.CommandLine pinned at 3.0.0-preview.4 (intentional - pre-release)
- Monthly Activity issue: June 2026 is #311
- PR #333 by NikiforovAll: Strict OOXML fix, all CI passing, needs maintainer merge decision
- Next: Continue WmlToXml.cs complex as+null patterns (ProduceXmlTransform, InjectComment mainPart blocks)
