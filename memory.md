# Clippy Memory

## Last Run
2026-07-15 16:09 UTC — Run 29431131026

## Comments Made
- #54: ManageMediaCopy CRC32 improvement idea
- #67: DocumentAssembler uses XPath 1.0, conditional rows workaround
- #77: Limitations of DocumentAssembler re: TOC/figure captions
- #103: DocumentAssembler uses content controls not bookmarks
- #28: ExcelAssembler design sketch + implementation follow-up (PR #432)
- #66: RegisterCustomHandler design + linked PR #417
- #380: Confirmed ArgumentNullException root cause in WmlComparer Hashing.cs
- #381: Confirmed IndexOutOfRangeException root cause in TextReplacer.cs
- #385: Confirmed InvalidCastException root cause in WmlComparer Lcs.cs
- #383: Confirmed rowspan/gridSpan root cause + linked fix PR
- #384: Confirmed hyperlink anchor root cause + linked fix PR
- #386: Confirmed multi-part root causes; parts 1+2+3 now fixed via PRs
- #374: word build implementation guidance
- #377: simplify-markup implementation guidance
- #401: linked NuGet OIDC PR (now merged #407)

## Open Clippy PRs
- #432: feat(excel): ExcelAssembler — closes #28 (squash rebase, clean)
- #aw_frozen_dicts (TBD number): perf: convert static readonly Dictionary to FrozenDictionary (Task 5, 2026-07-15)

## Backlog Cursor
Last issue processed: #401. All issues labelled. All issues have Clippy comments.

## Notes
- v3.7.0 released (CLI v0.7.0)
- PR #430 merged (StringExtensions tests + fr_FR dead code removal)
- PR #431 merged (deps: SkiaSharp 4.150.1, TUnit 1.59.0, Test.Sdk 18.8.1)
- PR #432 open (ExcelAssembler, supersedes #165)
- PR #433 merged (actions/setup-node from 6 to 7)
- Monthly Activity issue July 2026: #370
- SixLabors.ImageSharp.Drawing 3.0.0 requires paid commercial license — do NOT upgrade
- All deps current as of 2026-07-15
- Issue #386 closed (all parts fixed)
- Issue #28 open, has PR #432 awaiting merge

## Future Ideas
- Remaining non-frozen static dicts in Html/ (HtmlToWmlCssApplier ColorMap, FontSizeMap, BorderStyleMap)
- PresentationBuilder stream leaks from #386 part 4 (overlaps #25/#15)
