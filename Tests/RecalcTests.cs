using TVE.PureDocs.Excel;
using Xunit;

namespace PureDocs.Excel.Tests;

public class RecalcTests
{
    [Fact]
    public void FullRecalc_CachesAllFormulaCells()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetValue(10.0); s["A2"].SetValue(20.0);
        s["A3"].SetFormula("A1+A2");
        s["A4"].SetFormula("A3*2");
        s.SmartRecalculate();
        // Regression: GetRecalcOrderWithCycles previously omitted the changed formula cells.
        Assert.Equal(30, s.Cached("A3"));
        Assert.Equal(60, s.Cached("A4"));
    }

    [Fact]
    public void CellEdit_MarksDependentsDirty_AndIncrementalRecalcUpdates()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetValue(10.0); s["A2"].SetValue(20.0);
        s["A3"].SetFormula("A1+A2");
        s["A4"].SetFormula("A3*2");
        s.SmartRecalculate();

        s["A1"].SetValue(100.0);
        Assert.True(s.CalcChain.DirtyCount > 0);   // edit auto-marked dirty
        s.SmartRecalculate();
        Assert.Equal(120, s.Cached("A3"));
        Assert.Equal(240, s.Cached("A4"));         // transitive dependent
    }

    [Fact]
    public void EditingFormulaText_RecalculatesThatCell()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetValue(10.0);
        s["B1"].SetFormula("A1*2");
        s.SmartRecalculate();
        Assert.Equal(20, s.Cached("B1"));

        s["B1"].SetFormula("A1+1");
        s.SmartRecalculate();
        Assert.Equal(11, s.Cached("B1"));
    }

    [Fact]
    public void LongDependencyChain_PropagatesIncrementally()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetValue(1.0);
        for (int r = 2; r <= 10; r++) s[$"A{r}"].SetFormula($"A{r - 1}+1");
        s.SmartRecalculate();
        Assert.Equal(10, s.Cached("A10"));

        s["A1"].SetValue(100.0);
        Assert.Equal(9, s.SmartRecalculate());  // 9 dependent formula cells
        Assert.Equal(109, s.Cached("A10"));
    }

    [Fact]
    public void BulkSetValues_RecalculatesViaSuspendResume()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        for (int r = 1; r <= 5; r++) s[$"A{r}"].SetValue(r * 1.0);
        s["B1"].SetFormula("SUM(A1:A5)");
        s.SmartRecalculate();
        Assert.Equal(15, s.Cached("B1"));

        s.GetRange("A1:A5").SetValues(new object[,] { { 10.0 }, { 10.0 }, { 10.0 }, { 10.0 }, { 10.0 } });
        s.SmartRecalculate();
        Assert.Equal(50, s.Cached("B1"));
    }

    [Fact]
    public void CircularReference_ReportsRefError()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetFormula("A2+1");
        s["A2"].SetFormula("A1+1");
        s.SmartRecalculate();
        var v = s.CalcChain.GetCachedValue(TVE.PureDocs.Excel.Formulas.CellAddress.FromReference("A1"));
        Assert.True(v.HasValue && v.Value.IsError);
    }
}
