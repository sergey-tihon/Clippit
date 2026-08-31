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

## Run 2026-08-21 15:52 UTC (32499715020)
- Selected tasks: 2 (Issue Investigation and Comment), 4 (Engineering Investments), 3 (Issue Investigation and Fix)
- Task 4/3/6 combined: PR #449 (ExcelAssembler) had merge conflicts against current master (confirmed via `git merge --no-commit`: conflicts in actions-lock.json, clippy.lock.yml, Clippit.Cli.csproj, Clippit.Tests.csproj, FormattingAssembler.cs). Rebased branch clippy/improve-excel-assembler-rebase-20260725-2cc993590e0f6c6f onto master (fe779ac) - rebase applied cleanly with no manual conflict resolution needed (all prior master commits were already incorporated from earlier rebases). Build (0 errors), csharpier check, and full test suite (2357 tests, 2355 passed, 2 skipped) all pass. Pushed update via push_to_pull_request_branch.
- Task 4: dotnet outdated (2026-08-21) shows no outdated dependencies.
- Task 3: No new bug/help-wanted/good-first-issue labelled issues found beyond existing #67/#77/#103 (all Q&A, already answered, not "fixable" bugs). No fixable issues this run.
- Task 2: reviewed #67, #77, #103 - no new human activity since last Clippy comments, no re-engagement needed.
- Updated Monthly Activity issue #460 with new run entry.

## Run 2026-08-20 15:52 UTC (32388481883)
- Selected tasks: 5 (Coding Improvements), 1 (Labelling → substituted, no unlabelled issues), 4 (Engineering Investments)
- Task 4: dotnet outdated found TUnit 1.64.13->1.65.38, Microsoft.NET.Test.Sdk 18.8.1->18.9.0. Created draft PR (branch clippy/eng-deps-tunit-testsdk-20260820). Build+csharpier+2345 tests(2343 pass/2 skip) pass.
- Task 1 substituted with Task 2 review: no unlabelled issues exist; reviewed #67, #77, #103 for new human activity — none found since last Clippy comments, no re-engagement needed.
- Task 5: found and fixed remaining ContainsKey+indexer double-lookup pattern in Clippit/Word/RevisionProcessor.cs (3 occurrences: moveFromRangeEnd, customXmlDelRangeEnd, customXmlMoveFromRangeEnd) and Clippit/Comparer/WmlComparer.Private.Methods.Hashing.cs (1 occurrence, source/target Unid correlation). Created draft PR (branch clippy/improve-doublelookup-revisionprocessor-hashing-20260820). Build+csharpier+2345 tests(2343 pass/2 skip) pass.
- Confirmed PR #477 (HtmlToWmlCssApplier double-lookup fix from run 31895318701) was merged onto master (commit 85164d3).
- Updated Monthly Activity issue #460 with new run entry and refreshed suggested actions (added the two new PRs, kept comment-check items for #67/#77/#103).

## Run 2026-08-23 15:54 UTC (32649513366)
- Selected tasks: 3 (Issue Investigation and Fix), 2 (Issue Investigation and Comment), 5 (Coding Improvements)
- Task 3/6: PR #449 (ExcelAssembler) was behind master (merge-base 469bc56, current master 6380b15). Rebased cleanly onto master (git rebase, no conflicts). Build (0 errors), csharpier check pass, full test suite 2364 tests (2362 pass, 2 skip) pass. Pushed via push_to_pull_request_branch.
- Task 3: no new fixable bug/help-wanted/good-first-issue issues found beyond existing Q&A issues (#67/#77/#103).
- Task 2: reviewed #67, #77, #103 - no new human activity since last Clippy comments, no re-engagement needed.
- Task 5: substituted - folded into rebase work for PR #449 (no additional standalone coding improvement made this run given time spent on rebase).
- Updated Monthly Activity issue #460 with new run entry and refreshed suggested actions.

## Run 2026-08-22 15:46 UTC (32582616455)
- Selected tasks: 9 (Testing Improvements), 5 (Coding Improvements), 3 (Issue Investigation and Fix)
- Confirmed PR #482 (RevisionProcessor/Hashing double-lookup) and #481 (TUnit/Test.Sdk deps) merged since last run.
- Task 9/5 combined: reviewed PtUtil.cs coverage gaps vs PtUtilsTests.cs; added tests for AddElementIfMissing, DescendantsTrimmedBeforeSelfReverseDocumentOrder, GetXElement/GetXmlNode/GetXDocument/GetXmlDocument roundtrips, FileUtils.GetFilesRecursive, FileUtils.ThreadSafeCreateDirectory. Also found+fixed a remaining ContainsKey+indexer double-lookup in SmlToHtmlConverter.CreateFontCssProperty. Build+csharpier+2352 tests(2350 pass/2 skip) pass. Created draft PR clippy/test-and-improve-ptutils-smltohtml-20260822.
- Task 3: no fixable bug/help-wanted/good-first-issue issues found beyond existing Q&A issues (#67/#77/#103, all already answered).
- Task 2: reviewed #67,#77,#103 - no new human activity since last Clippy comments, no re-engagement needed.
- PR #449 (ExcelAssembler) still open, awaiting maintainer review.

## Run 2026-08-25 15:56 UTC (32868709719)
- Selected tasks: 2 (Issue Investigation and Comment), 5 (Coding Improvements), 6 (Maintain Clippy PRs)
- Task 6: PR #449 (ExcelAssembler) was behind master again. Rebased onto master (6380b15) - clean rebase, no conflicts. Build (0 errors), csharpier check pass, full suite 2364 tests (2362 pass/2 skip) pass. Pushed via push_to_pull_request_branch.
- Task 5: found and fixed remaining ContainsKey+indexer double-lookup pattern in Clippit/Excel/SmlToHtmlConverter.cs (WriteStyleSheet, PtStyleName lookup used 3x for same key) — replaced with TryGetValue. Build+csharpier+2364 tests pass. Created draft PR (branch clippy/improve-smltohtml-doublelookup-20260825).
- Task 2: reviewed #67, #77, #103 - no new human activity since last Clippy comments, no re-engagement needed.
- Updated Monthly Activity issue #460 with new run entry and refreshed suggested actions (added new draft PR entry).

## Run 2026-08-26 (32986769210)
- Task 2: reviewed #67,#77,#103 - no new human activity, no re-engagement
- Task 3: no fixable issues found
- Task 6/10 (substituted): PR #449 rebased onto master caba0be, build+csharpier+2364 tests(2362 pass/2 skip) pass, pushed

## Run 2026-08-27 (33107342352)
- Selected tasks: 2 (Issue Investigation and Comment), 9 (Testing Improvements → substituted with Task 5), 10 (Take Repo Forward → substituted with Task 5)
- Task 2: reviewed #67,#77,#103 - no new human activity since last Clippy comments, no re-engagement needed.
- Issue #490 checked: gh-aw Protected Files notice for PR #449 — only workflow/lockfile diffs (.github/aw/actions-lock.json, clippy.lock.yml, clippy.md) from earlier rebases against master, no CHANGELOG/README changes. No action needed (informational, not a task item); left for maintainer awareness.
- Task 5 (substituted for 9/10): searched remaining ContainsKey+indexer double-lookup patterns across Clippit/. Found and fixed one in Clippit/Word/DocumentBuilder.cs (CopyStylesAndFonts: styleNameMap[styleName] indexed twice for comparison + correctionList.Add). Build (0 errors) + csharpier + full suite (2352 tests, 2350 pass/2 skip) all pass. Created draft PR (branch clippy/improve-doublelookup-documentbuilder-20260827).
- PR #449 (ExcelAssembler) still open, awaiting maintainer review; last rebased 2026-08-26 onto caba0be.

## Run 2026-08-28 19:37 UTC (33204538276)
- Selected tasks: 5 (Coding Improvements), 4 (Engineering Investments), 3 (Issue Investigation and Fix)
- Task 4: dotnet outdated found TUnit 1.65.38 -> 1.65.68. Created draft PR (branch clippy/eng-deps-tunit-20260828). Build+csharpier+2352 tests(2350 pass/2 skip) pass.
- Task 5: searched for remaining ContainsKey+indexer double-lookup patterns across Clippit/ — none remain (all previously found instances already fixed in prior PRs). Substituted with rebasing PR #449.
- Task 6 (substituted): PR #449 (ExcelAssembler) was behind master. Rebased cleanly onto master (5548b1b1), no conflicts. Build (0 errors), csharpier check pass, full suite 2364 tests (2362 pass/2 skip) pass. Pushed via push_to_pull_request_branch.
- Task 3: no new fixable bug/help-wanted/good-first-issue issues found beyond existing Q&A issues (#67/#77/#103).
- Confirmed PRs #489 (SmlToHtmlConverter CSS) and #491 (DocumentBuilder style ID) merged since last run.
- Updated Monthly Activity issue #460 with new run entry and refreshed suggested actions.

## Run 2026-08-29 15:42 UTC (33260910953)
- Selected tasks: 2 (Issue Investigation and Comment), 3 (Issue Investigation and Fix), 5 (Coding Improvements)
- Task 6 (substituted, PR maintenance): PR #449 (ExcelAssembler) was behind master (merge-base 469bc56, master c747ebf). Rebased cleanly onto master (no conflicts). Build (0 errors) + csharpier + full suite 2364 tests (2362 pass/2 skip) all pass. Pushed via push_to_pull_request_branch using local branch renamed to match PR branch `clippy/improve-excel-assembler-rebase-20260725-2cc993590e0f6c6f`.
- Task 2: reviewed #67, #77, #103 - no new human activity since last Clippy comments, no re-engagement needed.
- Task 3: no new fixable bug/help-wanted/good-first-issue issues found.
- Task 5 substituted: folded into rebase work.
- **IMPORTANT BUG TO FIX NEXT RUN**: attempted to update Monthly Activity issue #460 via `printf '{...}' | safeoutputs update_issue .` but shell quoting mangled the JSON (nested single quotes in body text broke printf escaping), resulting in argumentBytes=2 (empty/near-empty payload) being sent successfully — this likely WIPED or corrupted issue #460's body since the tool reported success. update_issue has a 1-per-run limit so a retry this run was blocked ("E002: update_issue limit reached"). **Next run: check issue #460's body first thing — if it's empty/wrong, immediately fix it with a properly-escaped call using a heredoc file (like `/tmp/gh-aw/agent/issue460.json` via `safeoutputs update_issue . < file.json`), NOT printf with inline JSON containing apostrophes/emoji.**
- Lesson learned: always write safeoutputs JSON payloads to a temp file via heredoc (cat > file << 'EOF') and pipe with `< file`, never printf with inline complex strings.

## Run 2026-08-30 15:42 UTC (33320328645)
- Selected tasks: 10 (Take Repo Forward), 4 (Engineering Investments), 9 (Testing Improvements)
- Task 6 (substituted for 10, PR maintenance): PR #449 (ExcelAssembler) was 31 commits behind master (merge-base 469bc56, master c747ebf). Rebased cleanly onto master (no conflicts). Build (0 errors) + csharpier + full suite 2364 tests (2362 pass/2 skip) all pass. Pushed via push_to_pull_request_branch.
- Task 4: dotnet outdated (2026-08-30) shows no outdated dependencies (TUnit already bumped to 1.65.68 via merged PR #493).
- Task 9: reviewed Clippit/Word/Assembler/ coverage gap - AssemblerInternalsTests.cs covered XPathExtensions.EvaluateXPath/EvaluateXPathToString and ErrorHandler but not XElementExtensions (IsPlainText, MergeRunProperties) or UriExtensions.GetUri or XPathExtensions.TryEvalueStringToByteArray. Added 11 new unit tests. Build+csharpier+2375 tests(2373 pass/2 skip) pass. Created draft PR (branch clippy/test-assembler-xelement-uri-extensions-20260830).
- Verified last run's suspected issue #460 corruption did NOT actually happen - body was intact and well-formed; only needed a fresh run entry appended.
- Searched for remaining ContainsKey+indexer double-lookup patterns across Clippit/ - none found (confirmed again).
- Updated Monthly Activity issue #460 successfully via heredoc-file approach (`cat > file << 'EOF'` + python json validation + `safeoutputs update_issue . < file`) - confirms the lesson from last run's near-miss works correctly.

## Run 2026-08-31 15:43 UTC (33409933591)
- Selected tasks: 9 (Testing Improvements), 3 (Issue Investigation and Fix), 10 (Take Repo Forward)
- Confirmed PR #496 (XElementExtensions/UriExtensions tests) merged onto master (a11ae3e) since last run.
- Task 10 (substituted, PR maintenance): PR #449 (ExcelAssembler) was 32 commits behind master. Rebased cleanly onto master (a11ae3e), no conflicts. Build (0 errors) + csharpier + full suite 2375 tests (2373 pass/2 skip) all pass. Pushed via push_to_pull_request_branch.
- Task 9: found FileDataExtensions.GetBase64EncodedDocumentElement (internal DocumentAssembler helper, used for DocumentTemplate/Document placeholder resolution) had no direct unit tests - only indirect integration coverage. Added 2 unit tests in AssemblerInternalsTests.cs. Build+csharpier+2365 tests(2363 pass/2 skip) pass. Created draft PR (branch clippy/test-filedataextensions-20260831).
- Task 3: no new fixable bug/help-wanted/good-first-issue issues found beyond existing Q&A issues (#67/#77/#103, all already answered, no new human activity).
- Task 2 (implicit review): reviewed #67, #77, #103 - no new human activity since last Clippy comments, no re-engagement needed.
