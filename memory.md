# Clippy Memory

## Last Run
2026-09-25 15:56 UTC — Run 36156071876 (see memory.json for detailed structured state)
- Tasks selected: 8 (Performance Improvements), 9 (Testing Improvements), 10 (Take Repository Forward)
- Confirmed both prior draft PRs (clippy/eng-tunit-1.69.0-20260924 #530, clippy/test-htmltowmlconverter-cleanupcss-emu-20260924 #531) merged onto master since last run
- Issues #67/#77/#103: no new human activity since last Clippy comments - Task 2 not applicable; all 4 open issues already labelled - Task 1 fallback also n/a
- Task 8: delegated exploration found a genuine new O(n^2) hot-path bug in DocumentAssembler.cs TransformToMetadata - occurrence-counting of "<#" directive markers used paraContents.Select((_, i) => paraContents.Substring(i)).Count(sub => sub.StartsWith("<#")) which allocates O(n) substrings per call and runs on EVERY A.r/W.p element during template assembly (the hottest path in the templating engine). Fixed both occurrences (A.r and W.p branches) with a new CountTemplateDirectiveStarts helper using a single-pass IndexOf scan, zero allocations, identical behavior. Added regression test DA_Content_MultipleDirectivesInSingleRun_BothReplaced covering the occurrences!=1 multi-directive regex-replacement branch. build+csharpier+2462 tests pass (2460 pass/2 skip); created draft PR clippy/perf-documentassembler-occurrence-count-20260925
- Task 9/10 substituted: exploration for future testing candidates found XPathExtensions.cs (Clippit/Word/Assembler/) and FluentPresentationBuilder.Deduplication.cs (Clippit/PowerPoint/Fluent/) as next zero-direct-coverage candidates (only exercised indirectly via integration tests) - noted for a future run, not implemented this run since Task 8 finding was the highest-value work
- Rewrote Monthly Activity issue #501 Run History/Suggested Actions to reflect this run

## Previous Run
2026-09-24 16:04 UTC — Run 36022357929 (see memory.json for detailed structured state)
- Tasks selected: 2 (Issue Comment), 4 (Engineering Investments), 3 (Issue Fix)
- Confirmed both prior draft PRs (clippy/perf-where-firstordefault-lastordefault-20260922 #528, clippy/test-wmlcomparerextensions-20260923 #529) merged onto master since last run
- Issues #67/#77/#103: no new human activity since last Clippy comments - Task 2 not applicable
- Task 3: no fixable bug/help-wanted/good-first-issue issues found (only Q&A-style help-wanted threads) - substituted with Task 9 testing work
- Task 4: dotnet outdated showed only TUnit 1.68.17->1.69.0 (minor, test-only) - created draft PR clippy/eng-tunit-1.69.0-20260924; build+csharpier+2440 tests pass (2438 pass/2 skip)
- Task 5 fallback exploration: no new low-risk coding improvements found - ContainsKey double-lookups, Where().FirstOrDefault()/Any() chains, FrozenSet conversions, dead code all already swept in prior runs
- Task 9: HtmlToWmlConverter.cs CleanUpCss/Emu/TPoint/Twip (previously zero direct test coverage, pure helpers only exercised indirectly via integration tests) - added 15 unit tests in new Clippit.Tests/Html/HtmlToWmlConverterUnitTests.cs; build+csharpier+2461 tests(2459 pass/2 skip) pass; created draft PR clippy/test-htmltowmlconverter-cleanupcss-emu-20260924
- Rewrote Monthly Activity issue #501 Run History/Suggested Actions to reflect this run

## Previous Run
2026-09-23 15:47 UTC — Run 35883752713 (see memory.json for detailed structured state)
- Tasks selected: 9 (Testing Improvements), 5 (Coding Improvements), 2 (Issue Comment)
- Confirmed prior draft PR (clippy/perf-where-firstordefault-lastordefault-20260922) still open, no CI issues; no other open Clippy PRs otherwise
- Issues #67/#77/#103: no new human activity since last Clippy comments - Task 2 not applicable; Task 1 fallback also n/a (all 4 open issues already labelled)
- Task 9: WmlComparerExtensions (Clippit/Comparer/WmlComparerExtensions.cs) had zero test coverage - added 5 unit tests (GetMainDocumentRoot/GetMainDocumentBody happy+failure paths, GetXElement) in new Clippit.Tests/Comparer/WmlComparerExtensionsTests.cs; build+csharpier+2439 tests(2437 pass/2 skip) pass; created draft PR clippy/test-wmlcomparerextensions-20260923
- Task 5: dotnet outdated showed only TUnit 1.68.17->1.69.0 (minor, test-only) - not a standalone PR-worthy change this run; folded effort into Task 9 testing work
- Rewrote Monthly Activity issue #501 body from scratch (had accumulated duplicated sections from repeated append updates in prior runs)

## Previous Run
2026-09-21 15:43 UTC — Run 35620722528 (see memory.json for detailed structured state)
- Tasks selected: 3 (Issue Fix), 5 (Coding Improvements), 8 (Performance Improvements)
- Confirmed previous draft PR clippy/test-wmlcomparerutil-20260920 merged onto master as #525; no open Clippy PRs at run start; dotnet outdated shows no outdated deps
- Issues #67/#77/#103: no new human activity since last Clippy comments - Task 2 not applicable
- Task 3: no fixable bug/help-wanted/good-first-issue issues found beyond existing thoroughly-answered question threads - substituted with Task 5/8 work
- Task 5/8: FormattingAssembler.CharStyleAttributes had duplicate nested TogglePropertyNames/PropertyNames arrays shadowing identical outer-class fields (dead code) - removed; converted FormattingAssembler.TogglePropertyNames and RevisionProcessor.BlockLevelElements (Contains()-only usage) to FrozenSet<XName>; fixed ContainsKey+indexer double lookup on styleNameMap in DocumentBuilder.WriteStylesXml; build+csharpier+2434 tests pass (2432 pass/2 skip); created draft PR clippy/improve-frozenset-duplicate-cleanup-20260921
- Rewrote Monthly Activity issue #501 Run History/Suggested Actions to reflect this run

## Previous Run
2026-09-20 15:43 UTC — Run 35520367180 (see memory.json for detailed structured state)
- Tasks selected: 9 (Testing Improvements), 8 (Performance Improvements), 2 (Issue Comment)
- Confirmed all 3 previously-tracked draft PRs (SkiaSharp/TUnit deps #524, MhtParser OrdinalIgnoreCase fix #523, FrozenSet perf #522) merged onto master; no open PRs at run start
- Issues #67/#77/#103: no new human activity since last Clippy comments - Task 2 not applicable
- Task 9: WmlComparerUtil (Clippit/Comparer/WmlComparerUtil.cs) had zero test coverage - added 16 unit tests (hex encoding, SHA1 hashing stackalloc/heap paths, local-name to ComparisonUnitGroupType mapping incl. error case) in new Clippit.Tests/Comparer/WmlComparerUtilTests.cs; build+csharpier+2433 tests pass (2431 pass/2 skip); created draft PR clippy/test-wmlcomparerutil-20260920
- Task 8: no additional clear performance opportunity found beyond the testing work this run; folded into Task 9
- Monthly Activity issue #501 body had reverted to stale/duplicated state (old PR #449, deleted issues #505/#507) despite prior cleanup runs - rewrote from scratch again

## Previous Run
2026-09-19 15:42 UTC — Run 35452566842 (see memory.json for detailed structured state)
- Tasks selected: 2 (Issue Comment), 4 (Engineering Investments), 5 (Coding Improvements)
- Issues #67/#77/#103: no new human activity since last Clippy comments - Task 2 not applicable; no unlabelled issues - Task 1 fallback also n/a
- Task 4: dotnet outdated found SkiaSharp 4.152.0->4.152.1 (+ Linux native assets pkg), TUnit 1.68.4->1.68.17 (both patch) - created draft PR clippy/eng-deps-skiasharp-tunit-20260919; build+csharpier+2414 tests pass
- Task 5: found MhtParser.Parse (PtUtil.cs) had two StartsWith("boundary")/StartsWith("charset") calls missing StringComparison.OrdinalIgnoreCase, inconsistent with every other StartsWith in the same method - fixed both + added regression test PU002_MixedCaseBoundaryAndCharSetAreParsed; created draft PR clippy/improve-mhtparser-stringcomparison-20260919; build+csharpier+2415 tests pass
- Discovered PR #449 (ExcelAssembler) is CLOSED not merged (previously tracked as open); issues #505/#507 (gh-aw Protected Files notices) were DELETED (410) upstream - removed all three stale entries from Monthly Activity issue #501
- Rewrote Monthly Activity issue #501 to remove stale/deleted references and add this run's new PRs
