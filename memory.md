# Clippy Memory

## Last Run
2026-06-30 16:30 UTC — Run 28459820592

## Comments Made
- #54: CRC32 improvement idea — explained reflection is fragile; suggested profiling first
- #67: Explained DocumentAssembler uses XPath 1.0; pointed to conditional row workaround
- #77: Explained limitations of DocumentAssembler re: figure captions/TOC update
- #103: Explained DocumentAssembler uses content controls not bookmarks
- #28: Design sketch for ExcelAssembler + follow-up linking to PR #165
- #66: Explained RegisterCustomHandler design + linked PR #163
- #364: Noted PR #365 (word compare CLI) created by Copilot coding agent

## Open Clippy PRs
- #163: feat: RegisterCustomHandler extensibility API — 131+ commits behind master
- #165: feat(excel): ExcelAssembler — 131+ commits behind master

## Other Notable PRs
- #365: Add `word compare` CLI command (by Copilot coding agent) — addresses issue #364, draft, pending review

## Backlog Cursor
All 7 non-activity open issues (#28, #54, #66, #67, #77, #103, #364) have Clippy comments.
Issue #364 is the newest; Copilot SWE agent already created PR #365 for it.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (2026-06-23)
- Monthly Activity issue: June 2026 is #311
- Null modernization series COMPLETE for all core library files
- FrozenSet PR was #361 (merged 2026-06-29), FrozenDict PR was #362 (merged 2026-06-29)
  (Previous memory incorrectly expected #363/#364; those were npm ARM64 PRs by maintainer)
- SixLabors.ImageSharp.Drawing 3.0.0 requires paid commercial license — do NOT upgrade

## Future Ideas
- ExcelAssembler row-repetition support (follow-up to #165)
- PRs #163/#165: 131+ commits behind master — maintainer should rebase or close
