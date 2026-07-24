using TVE.PureDocs.Excel;
using Xunit;

namespace PureDocs.Excel.Tests;

public class NamedRangeTests
{
    private static Workbook TwoSheetBook(out Worksheet s1, out Worksheet s2)
    {
        var wb = Workbook.Create();
        s1 = wb.AddWorksheet("Sheet1");
        s2 = wb.AddWorksheet("Sheet2");
        for (int r = 1; r <= 5; r++) s1[$"A{r}"].SetValue(r * 10.0); // 10..50
        s2["A1"].SetValue(999.0);
        return wb;
    }

    [Fact]
    public void WorkbookScopedName_ResolvesInFormula()
    {
        using var wb = TwoSheetBook(out var s1, out _);
        wb.NamedRanges.Define("GiaBan", "Sheet1!$C$2:$C$6");
        for (int r = 2; r <= 6; r++) s1[$"C{r}"].SetValue(r * 10.0);
        Assert.Equal(200, s1.Num("SUM(GiaBan)"));
    }

    [Fact]
    public void SingleCellName_Resolves()
    {
        using var wb = TwoSheetBook(out var s1, out _);
        wb.NamedRanges.Define("Rate", "Sheet1!$A$1");
        Assert.Equal(20, s1.Num("Rate*2"));
    }

    [Fact]
    public void NameLookup_IsCaseInsensitive()
    {
        using var wb = TwoSheetBook(out var s1, out _);
        wb.NamedRanges.Define("Rng", "Sheet1!$A$1:$A$5");
        Assert.Equal(150, s1.Num("SUM(rng)"));
    }

    [Fact]
    public void UndefinedName_ReturnsNameError()
    {
        using var wb = TwoSheetBook(out var s1, out _);
        Assert.Equal("#NAME?", s1.Str("SUM(Nope)"));
    }

    [Fact]
    public void SheetScopedNames_CoexistWithoutCollision()
    {
        using var wb = TwoSheetBook(out var s1, out var s2);
        wb.NamedRanges.Define("Local", "Sheet1!$A$1", 0); // scope Sheet1 -> 10
        wb.NamedRanges.Define("Local", "Sheet2!$A$1", 1); // scope Sheet2 -> 999
        Assert.Equal(10, s1.Num("Local"));
        Assert.Equal(999, s2.Num("Local"));
    }

    [Fact]
    public void ScopedName_WinsOverGlobal()
    {
        using var wb = TwoSheetBook(out var s1, out var s2);
        wb.NamedRanges.Define("Mix", "Sheet2!$A$1");    // global -> 999
        wb.NamedRanges.Define("Mix", "Sheet1!$A$1", 0); // Sheet1-scoped -> 10
        Assert.Equal(10, s1.Num("Mix"));   // scoped wins on Sheet1
        Assert.Equal(999, s2.Num("Mix"));  // global on Sheet2
    }

    [Fact]
    public void NamedRanges_RoundTripThroughFile()
    {
        var path = TestHelpers.TempXlsx();
        try
        {
            using (var wb = TwoSheetBook(out var s1, out _))
            {
                wb.NamedRanges.Define("Rng", "Sheet1!$A$1:$A$5");
                wb.NamedRanges.Define("Local", "Sheet1!$A$1", 0);
                wb.SaveAs(path);
            }
            using var wb2 = Workbook.Open(path);
            Assert.True(wb2.NamedRanges.IsDefined("Rng"));
            Assert.Equal(150, wb2.Worksheets["Sheet1"].Num("SUM(Rng)"));
            Assert.Equal(10, wb2.Worksheets["Sheet1"].Num("Local"));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
