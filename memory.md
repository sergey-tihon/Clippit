# Clippy Memory

## Last Run
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
