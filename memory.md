# Clippy Memory

## Last Run
2026-08-13 20:51 UTC — Run 31742767027 (see memory.json for detailed structured state - this file is legacy/secondary)

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
- #449: feat(excel): ExcelAssembler — rebased 2026-08-11 onto master (90cfe11), conflicts resolved, build+2317 tests pass, awaiting review (mergeable_state=dirty as of 2026-08-13, needs future rebase)
- #470: refactor: eliminate ContainsKey+indexer double lookups — merged
- #472: perf: eliminate double dictionary lookup in PtBucketTimer.Bucket — merged
- (draft) test: add unit tests for ColorParser (branch clippy/test-colorparser-20260813), build+csharpier+2313 tests pass, awaiting review
- (draft) perf: eliminate double dictionary/SortedList lookups in WorksheetAccessor and FieldRetriever (branch clippy/perf-double-lookups-worksheetaccessor-fieldretriever-20260813), build+csharpier+2307 tests pass, awaiting review

## Other Open PRs
- #474: alexeysp11's numbering format enum refactor (breaking change, not a Clippy PR), CI in progress as of 2026-08-13

## Backlog Cursor
Last issue processed: #401. All issues labelled. All issues have Clippy comments.

## Notes
- dotnet outdated (2026-08-12): no outdated dependencies
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

## Run 2026-08-15 16:31 UTC (31895318701)
- Selected tasks: 5 (Coding Improvements), 9 (Testing Improvements → substituted), 10 (Take Repo Forward → substituted)
- Task 5: Found remaining ContainsKey+indexer double-lookup pattern in HtmlToWmlCssApplier.cs (AddPropertyToElement, AddPropertyToDictionary) — hot path called per CSS property per element. Fixed with TryGetValue. Build+csharpier+2313 tests pass. Created draft PR (branch clippy/perf-htmlcssapplier-double-lookup-20260815).
- Task 9/10 substituted with Task 5 work since the perf/coding-improvement candidate found was higher value than speculative test additions or forward-looking exploration this run.
- Confirmed via memory review that Html/ FrozenDictionary maps (ColorMap, FontSizeMap, BorderStyleMap) were already converted in earlier runs — updated stale "Future Ideas" note.
- Updated Monthly Activity issue #460 with new run entry and refreshed suggested actions (removed merged PR entries #470/#472/#473/#475 references, added new draft PR).
