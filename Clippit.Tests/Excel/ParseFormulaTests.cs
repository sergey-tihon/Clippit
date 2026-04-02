// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using Clippit.Excel;

namespace Clippit.Tests.Excel;

/// <summary>
/// Unit tests for <see cref="ParseFormula"/>, which parses Excel formula strings
/// (via a PEG grammar) and provides helpers to:
/// <list type="bullet">
/// <item><description><see cref="ParseFormula.ReplaceSheetName"/> — rename sheet references within a formula</description></item>
/// <item><description><see cref="ParseFormula.ReplaceRelativeCell"/> — adjust relative cell references by a row/column offset</description></item>
/// </list>
/// </summary>
public class ParseFormulaTests
{
    // ── ReplaceSheetName ────────────────────────────────────────────────────

    [Test]
    public async Task PF001_ReplaceSheetName_SimpleReference_Renamed()
    {
        var result = new ParseFormula("Sheet1!A1").ReplaceSheetName("Sheet1", "Sheet2");
        await Assert.That(result).IsEqualTo("Sheet2!A1");
    }

    [Test]
    public async Task PF002_ReplaceSheetName_NoSheetReference_Unchanged()
    {
        var result = new ParseFormula("A1+B2").ReplaceSheetName("Sheet1", "Sheet2");
        await Assert.That(result).IsEqualTo("A1+B2");
    }

    [Test]
    public async Task PF003_ReplaceSheetName_NonMatchingSheet_Unchanged()
    {
        var result = new ParseFormula("Sheet1!A1").ReplaceSheetName("Sheet2", "Sheet3");
        await Assert.That(result).IsEqualTo("Sheet1!A1");
    }

    [Test]
    public async Task PF004_ReplaceSheetName_MultipleOccurrences_AllReplaced()
    {
        var result = new ParseFormula("Sheet1!A1+Sheet1!B2").ReplaceSheetName("Sheet1", "Data");
        await Assert.That(result).IsEqualTo("Data!A1+Data!B2");
    }

    [Test]
    public async Task PF005_ReplaceSheetName_RangeReference_BothEndsRenamed()
    {
        // Single-sheet area reference "Sheet1!A1:B5"
        var result = new ParseFormula("Sheet1!A1:B5").ReplaceSheetName("Sheet1", "Summary");
        await Assert.That(result).IsEqualTo("Summary!A1:B5");
    }

    [Test]
    public async Task PF006_ReplaceSheetName_InsideSumFunction_Renamed()
    {
        var result = new ParseFormula("SUM(Sheet1!A1:A10)").ReplaceSheetName("Sheet1", "Sheet2");
        await Assert.That(result).IsEqualTo("SUM(Sheet2!A1:A10)");
    }

    // ── ReplaceRelativeCell ─────────────────────────────────────────────────

    [Test]
    public async Task PF011_ReplaceRelativeCell_BothRelative_BothAdjusted()
    {
        // A1 + rowOffset=1, colOffset=1 → B2
        var result = new ParseFormula("A1").ReplaceRelativeCell(1, 1);
        await Assert.That(result).IsEqualTo("B2");
    }

    [Test]
    public async Task PF012_ReplaceRelativeCell_ZeroOffsets_Unchanged()
    {
        var result = new ParseFormula("B3").ReplaceRelativeCell(0, 0);
        await Assert.That(result).IsEqualTo("B3");
    }

    [Test]
    public async Task PF013_ReplaceRelativeCell_ColumnOverflow_ZtoAA()
    {
        // Z1 + colOffset=1 → AA; row 1 + rowOffset=1 → 2
        var result = new ParseFormula("Z1").ReplaceRelativeCell(1, 1);
        await Assert.That(result).IsEqualTo("AA2");
    }

    [Test]
    public async Task PF014_ReplaceRelativeCell_AbsoluteColumn_ColumnPreserved()
    {
        // $A1: column is absolute → not adjusted; row is relative → adjusted
        var result = new ParseFormula("$A1").ReplaceRelativeCell(2, 2);
        await Assert.That(result).IsEqualTo("$A3");
    }

    [Test]
    public async Task PF015_ReplaceRelativeCell_AbsoluteRow_RowPreserved()
    {
        // A$1: column is relative → adjusted; row is absolute → not adjusted
        var result = new ParseFormula("A$1").ReplaceRelativeCell(2, 2);
        await Assert.That(result).IsEqualTo("C$1");
    }

    [Test]
    public async Task PF016_ReplaceRelativeCell_BothAbsolute_Unchanged()
    {
        // $A$1: both absolute → no change
        var result = new ParseFormula("$A$1").ReplaceRelativeCell(1, 1);
        await Assert.That(result).IsEqualTo("$A$1");
    }

    [Test]
    public async Task PF017_ReplaceRelativeCell_AreaReference_AllRelativeRefsAdjusted()
    {
        // A1:B2 + rowOffset=1, colOffset=1 → B2:C3
        var result = new ParseFormula("A1:B2").ReplaceRelativeCell(1, 1);
        await Assert.That(result).IsEqualTo("B2:C3");
    }

    [Test]
    public async Task PF018_ReplaceRelativeCell_SumWithRange_RowsAdjusted()
    {
        // SUM(A1:A5) + rowOffset=3, colOffset=0 → SUM(A4:A8)
        var result = new ParseFormula("SUM(A1:A5)").ReplaceRelativeCell(3, 0);
        await Assert.That(result).IsEqualTo("SUM(A4:A8)");
    }

    [Test]
    public async Task PF019_ReplaceRelativeCell_NegativeOffset_DecrementsBothAxes()
    {
        // C3 + rowOffset=-1, colOffset=-1 → B2
        var result = new ParseFormula("C3").ReplaceRelativeCell(-1, -1);
        await Assert.That(result).IsEqualTo("B2");
    }

    [Test]
    public async Task PF020_ReplaceRelativeCell_MultiCellFormula_AllRelativeRefsAdjusted()
    {
        // A1+B2 with (+1, +1) → B2+C3
        var result = new ParseFormula("A1+B2").ReplaceRelativeCell(1, 1);
        await Assert.That(result).IsEqualTo("B2+C3");
    }
}
