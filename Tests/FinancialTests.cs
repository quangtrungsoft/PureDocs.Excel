using TVE.PureDocs.Excel;
using Xunit;

namespace PureDocs.Excel.Tests;

public class FinancialTests
{
    private static Worksheet Sheet(out Workbook wb)
    {
        wb = Workbook.Create();
        return wb.AddWorksheet("S");
    }

    [Fact]
    public void Pv_LumpSum_MatchesClosedForm()
    {
        var s = Sheet(out var wb);
        using (wb) Assert.Equal(-1000 / Math.Pow(1.1, 10), s.Num("PV(0.1,10,0,1000)"), 6);
    }

    [Fact]
    public void Fv_LumpSum_MatchesClosedForm()
    {
        var s = Sheet(out var wb);
        using (wb) Assert.Equal(1000 * Math.Pow(1.1, 10), s.Num("FV(0.1,10,0,-1000)"), 6);
    }

    [Fact]
    public void Pmt_ThenFv_PaysOffLoan()
    {
        var s = Sheet(out var wb);
        using (wb)
        {
            double pmt = s.Num("PMT(0.05/12,60,-10000)");
            Assert.True(pmt > 0);
            s["Z1"].SetValue(pmt);
            Assert.Equal(0, s.Num("FV(0.05/12,60,Z1,-10000)"), 4);   // loan fully repaid
            Assert.Equal(-10000, s.Num("PV(0.05/12,60,Z1)"), 4);     // recovers principal (outflow)
        }
    }

    [Fact]
    public void Pv_AnnuityDue_IsOrdinaryTimesOnePlusRate()
    {
        var s = Sheet(out var wb);
        using (wb)
            Assert.Equal(s.Num("PV(0.1,5,-100,0,0)") * 1.1, s.Num("PV(0.1,5,-100,0,1)"), 6);
    }

    [Fact]
    public void Npv_DiscountsCashFlows()
    {
        var s = Sheet(out var wb);
        using (wb)
        {
            Assert.Equal(100, s.Num("NPV(0.1,110)"), 9);
            Assert.Equal(200, s.Num("NPV(0.1,110,121)"), 9);
        }
    }

    [Fact]
    public void Irr_ZeroesNpv()
    {
        var s = Sheet(out var wb);
        using (wb)
        {
            double[] cf = { -100, 30, 40, 50, 60 };
            for (int i = 0; i < cf.Length; i++) s[$"A{i + 1}"].SetValue(cf[i]);
            double irr = s.Num("IRR(A1:A5)");
            double npv = 0;
            for (int t = 0; t < cf.Length; t++) npv += cf[t] / Math.Pow(1 + irr, t);
            Assert.True(Math.Abs(npv) < 1e-4, $"NPV at IRR should be ~0 but was {npv}");
        }
    }
}
