# Clippy Memory

## Last Run
2026-07-03 07:26 UTC — Run 28645307730

## Comments Made
- #54: ManageMediaCopy CRC32 improvement idea
- #67: DocumentAssembler uses XPath 1.0, conditional rows workaround
- #77: Limitations of DocumentAssembler re: TOC/figure captions
- #103: DocumentAssembler uses content controls not bookmarks
- #28: ExcelAssembler design sketch + linked PR #165
- #66: RegisterCustomHandler design + linked PR #163
- #380: Confirmed ArgumentNullException root cause in WmlComparer Hashing.cs (2026-07-03)
- #381: Confirmed IndexOutOfRangeException root cause in TextReplacer.cs (2026-07-03)
- #385: Confirmed InvalidCastException root cause in WmlComparer Lcs.cs (2026-07-03)

## Open Clippy PRs
- #163: feat: RegisterCustomHandler extensibility API — 131+ commits behind master
- #165: feat(excel): ExcelAssembler — 131+ commits behind master
- (TBD): fix TextReplacer/WmlComparer crashes (closes #381, #380, #385) — created 2026-07-03
- (TBD): docs v3.6.0 changelog — created 2026-07-03

## Backlog Cursor
Last issue processed: #390. All issues have been labelled.

## Notes
- v3.5.1 released: SkiaSharp 4.x, System.CommandLine preview.5, MS.NET.Test.Sdk 18.7.0 (2026-06-23)
- v3.6.0 release PR created 2026-07-03 (CHANGELOG only, docs PR)
- Monthly Activity issue July 2026: #370
- 18 issues labelled 2026-07-03: #373-#390 (CLI features + upstream port bugs)
- SixLabors.ImageSharp.Drawing 3.0.0 requires paid commercial license — do NOT upgrade
- Issues #373 and #390 are Master tracking issues; #374-#379 are CLI sub-issues; #380-#389 are upstream port bugs

## Future Ideas
- Port issues from #390: #382 (Optional on Conditional), #383 (rowspan fix), #384 (bookmark anchors), #386 (DocumentAssembler whitespace), #387 (HeaderRowCount), #388 (ExtendedPart chart), #389 (German/Spanish locale)
- CLI sub-issues: #374 (word build), #375 (word assemble), #376 (word accept-revisions), #377 (word simplify-markup), #378 (excel create), #379 (word consolidate)
- PRs #163/#165: 131+ commits behind master — maintainer should rebase or close
