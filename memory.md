# Clippy Memory

## Last Run
2026-07-05 16:02 UTC — Run 28746586082

## Comments Made
- #54: ManageMediaCopy CRC32 improvement idea
- #67: DocumentAssembler uses XPath 1.0, conditional rows workaround
- #77: Limitations of DocumentAssembler re: TOC/figure captions
- #103: DocumentAssembler uses content controls not bookmarks
- #28: ExcelAssembler design sketch + linked PR #165
- #66: RegisterCustomHandler design + linked PR #163
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
- #163: feat: RegisterCustomHandler extensibility API — 131+ commits behind master
- #165: feat(excel): ExcelAssembler — 131+ commits behind master
- #405: feat(cli): word assemble command — CI green ✅, closes #375
- (TBD): fix(word): UnicodeMapper.RunToString xml:space — closes #386 sub-fix 3 (created 2026-07-05)
- (TBD): feat(cli): word simplify-markup command — closes #377 (created 2026-07-05)

## Backlog Cursor
Last issue processed: #401. All issues labelled.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0
- v3.6.0 changelog PR #393 merged 2026-07-03
- Issues #376 (accept-revisions) and #401 (NuGet OIDC) closed by maintainer 2026-07-05
- PR #406 (accept-revisions) merged 2026-07-05
- PR #407 (NuGet OIDC) merged 2026-07-05
- PR #408 (DocumentAssembler xml:space) merged 2026-07-05 by sergey-tihon (part 1 of #386)
- PR #409 open (sergey-tihon's OpenXmlRegex lastRenderedPageBreak fix, part 2 of #386)
- Monthly Activity issue July 2026: #370
- SixLabors.ImageSharp.Drawing 3.0.0 requires paid commercial license — do NOT upgrade
- All deps current as of 2026-07-05
- Issue #386 still open: part 3 addressed by new PR, part 4 (stream leaks) overlaps #25/#15

## Future Ideas
- CLI: #378 (excel create), #379 (word consolidate), #374 (word build)
- PresentationBuilder stream leaks from #386 part 4 (overlaps #25/#15)
- PRs #163/#165: 131+ commits behind master — maintainer should rebase or close
