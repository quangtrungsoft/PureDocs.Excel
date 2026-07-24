using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using TVE.PureDocs.Excel;
using TVE.PureDocs.Excel.Formulas;
using Xunit;

namespace PureDocs.Excel.Tests;

public class SpillTests
{
    [Fact]
    public void OneDimensionalSpill_AnchorAndSpilledCells()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetFormula("SEQUENCE(3)");
        s.SmartRecalculate();
        Assert.Equal(1, s.Cached("A1"));          // anchor = top-left
        Assert.Equal(2, s.Spill("A2"));
        Assert.Equal(3, s.Spill("A3"));
        Assert.Equal(1, s.CalcChain.Spills.RegionCount);
    }

    [Fact]
    public void TwoDimensionalSpill()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["E1"].SetFormula("SEQUENCE(2,2)");
        s.SmartRecalculate();
        Assert.Equal(1, s.Cached("E1"));
        Assert.Equal(2, s.Spill("F1"));
        Assert.Equal(3, s.Spill("E2"));
        Assert.Equal(4, s.Spill("F2"));
    }

    [Fact]
    public void SpilledCells_ReadableFromOtherFormulas()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetFormula("SEQUENCE(3)");
        s["C1"].SetFormula("A2+A3");
        s.SmartRecalculate();
        Assert.Equal(5, s.Cached("C1"));
    }

    [Fact]
    public void BlockedSpill_ReturnsSpillError()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetFormula("SEQUENCE(3)");
        s["A2"].SetValue(99.0); // blocks the spill target
        s.SmartRecalculate();
        var a1 = s.CalcChain.GetCachedValue(CellAddress.FromReference("A1"));
        Assert.True(a1.HasValue && a1.Value.IsError);
        Assert.Equal("#SPILL!", a1!.Value.ToString());
    }

    [Fact]
    public void SpillShrinks_WhenFormulaChanges()
    {
        using var wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetFormula("SEQUENCE(3)");
        s.SmartRecalculate();
        Assert.Equal(3, s.Spill("A3"));

        s["A1"].SetFormula("SEQUENCE(2)");
        s.SmartRecalculate();
        Assert.Null(s.Spill("A3"));      // A3 no longer part of the region
        Assert.Equal(2, s.Spill("A2"));
    }

    [Fact]
    public void SavedFile_IsSchemaValid_AndHasArrayFormula()
    {
        var path = TestHelpers.TempXlsx();
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddWorksheet("S");
                s["A1"].SetFormula("SEQUENCE(3)");
                s.SmartRecalculate();
                wb.SaveAs(path);
            }

            using var doc = SpreadsheetDocument.Open(path, false);
            var errors = new OpenXmlValidator().Validate(doc).ToList();
            Assert.Empty(errors);

            // metadata part present (dynamic-array XLDAPR)
            Assert.Contains(doc.WorkbookPart!.GetPartsOfType<CellMetadataPart>(), _ => true);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void SpilledValues_PersistAndReopenWithoutFalseSpillError()
    {
        var path = TestHelpers.TempXlsx();
        try
        {
            using (var wb = Workbook.Create())
            {
                var s = wb.AddWorksheet("S");
                s["A1"].SetFormula("SEQUENCE(3)");
                s.SmartRecalculate();
                wb.SaveAs(path);
            }

            using var wb2 = Workbook.Open(path);
            var s2 = wb2.Worksheets["S"];
            Assert.Equal(2, s2.Num("A2"));   // spilled value persisted

            s2.SmartRecalculate();
            var a1 = s2.CalcChain.GetCachedValue(CellAddress.FromReference("A1"));
            Assert.True(a1.HasValue && !a1.Value.IsError);   // no false #SPILL! on reload
            Assert.Equal(1, a1!.Value.NumberValue);
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }
}
