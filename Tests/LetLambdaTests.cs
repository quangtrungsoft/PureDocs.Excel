using TVE.PureDocs.Excel;
using Xunit;

namespace PureDocs.Excel.Tests;

public class LetLambdaTests
{
    private static Worksheet Sheet(out Workbook wb)
    {
        wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        s["A1"].SetValue(10.0);
        s["A2"].SetValue(20.0);
        return s;
    }

    [Theory]
    [InlineData("LET(x,5,x+1)", 6)]
    [InlineData("LET(x,5,y,10,x+y)", 15)]
    [InlineData("LET(x,5,y,x*2,y+1)", 11)]          // later binding references earlier
    [InlineData("LET(s,A1+A2,s*2)", 60)]            // binding over cell refs
    [InlineData("LET(r,A1,SUM(r,A2))", 30)]
    public void Let(string formula, double expected)
    {
        var s = Sheet(out var wb);
        using (wb) Assert.Equal(expected, s.Num(formula), 9);
    }

    [Theory]
    [InlineData("LAMBDA(x,x+1)(5)", 6)]
    [InlineData("LAMBDA(x,y,x*y)(3,4)", 12)]
    [InlineData("LET(f,LAMBDA(x,x*x),f(4))", 16)]              // lambda bound via LET
    [InlineData("LET(n,10,g,LAMBDA(x,x+n),g(5))", 15)]         // closure over LET binding
    public void Lambda(string formula, double expected)
    {
        var s = Sheet(out var wb);
        using (wb) Assert.Equal(expected, s.Num(formula), 9);
    }

    [Fact]
    public void Lambda_ArityMismatch_IsValueError()
    {
        var s = Sheet(out var wb);
        using (wb) Assert.Equal("#VALUE!", s.Str("LAMBDA(x,y,x+y)(1)"));
    }

    [Fact]
    public void NamedRangesAndPercent_StillWorkAlongsideScope()
    {
        var s = Sheet(out var wb);
        using (wb)
        {
            wb.NamedRanges.Define("NR", "S!$A$1");
            Assert.Equal(30, s.Num("NR+A2"));
            Assert.Equal(0.5, s.Num("50%"), 9);
        }
    }
}
