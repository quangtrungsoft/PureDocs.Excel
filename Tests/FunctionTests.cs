using TVE.PureDocs.Excel;
using Xunit;

namespace PureDocs.Excel.Tests;

public class FunctionTests
{
    private static Worksheet Grid(out Workbook wb)
    {
        wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        // region A=amount, B=region label, C=qty
        string[] reg = { "N", "S", "N", "S", "N" };
        double[] amt = { 10, 20, 30, 40, 50 };
        double[] qty = { 1, 2, 3, 4, 5 };
        for (int i = 0; i < 5; i++)
        {
            s[$"A{i + 1}"].SetValue(amt[i]);
            s[$"B{i + 1}"].SetValue(reg[i]);
            s[$"C{i + 1}"].SetValue(qty[i]);
        }
        return s;
    }

    [Theory]
    [InlineData("SUMIFS(A1:A5,B1:B5,\"N\")", 90)]
    [InlineData("SUMIFS(A1:A5,B1:B5,\"N\",C1:C5,\">1\")", 80)]
    [InlineData("COUNTIFS(B1:B5,\"N\")", 3)]
    [InlineData("COUNTIFS(B1:B5,\"N\",A1:A5,\">=30\")", 2)]
    [InlineData("AVERAGEIFS(A1:A5,B1:B5,\"N\")", 30)]
    public void MultiConditionAggregates(string formula, double expected)
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal(expected, s.Num(formula), 9);
    }

    [Theory]
    [InlineData("SIN(PI()/2)", 1)]
    [InlineData("COS(0)", 1)]
    [InlineData("DEGREES(PI())", 180)]
    [InlineData("ATAN2(1,1)", Math.PI / 4)]
    [InlineData("SUMSQ(3,4)", 25)]
    public void TrigAndMath(string formula, double expected)
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal(expected, s.Num(formula), 9);
    }

    [Fact]
    public void RadiansRoundTrip()
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal(Math.PI, s.Num("RADIANS(180)"), 9);
    }

    [Fact]
    public void Asin_OutOfDomain_IsNum()
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal("#NUM!", s.Str("ASIN(2)"));
    }

    [Theory]
    [InlineData("LOOKUP(30,A1:A5,B1:B5)", "N")]
    [InlineData("TEXTBEFORE(\"hello-world\",\"-\")", "hello")]
    [InlineData("TEXTAFTER(\"hello-world\",\"-\")", "world")]
    [InlineData("TEXTAFTER(\"a.b.c\",\".\",2)", "c")]
    [InlineData("ADDRESS(2,3)", "$C$2")]
    [InlineData("ADDRESS(2,3,4)", "C2")]
    [InlineData("ADDRESS(1,1,1,FALSE)", "R1C1")]
    public void TextAndLookupText(string formula, string expected)
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal(expected, s.Str(formula));
    }

    [Theory]
    [InlineData("XLOOKUP(\"S\",B1:B5,A1:A5)", 20)]      // first S -> 20
    [InlineData("SUBTOTAL(9,A1:A5)", 150)]              // SUM
    [InlineData("SUBTOTAL(1,A1:A5)", 30)]               // AVERAGE
    [InlineData("SUBTOTAL(102,A1:A5)", 5)]              // COUNT (ignore-hidden variant)
    public void LookupNumeric(string formula, double expected)
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal(expected, s.Num(formula), 9);
    }

    [Fact]
    public void XLookup_NotFound_UsesIfNotFound()
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal("none", s.Str("XLOOKUP(\"Z\",B1:B5,A1:A5,\"none\")"));
    }

    [Theory]
    [InlineData("INDIRECT(\"A3\")", 30)]
    [InlineData("INDIRECT(\"A\"&3)", 30)]
    [InlineData("SUM(INDIRECT(\"A1:A3\"))", 60)]
    public void Indirect(string formula, double expected)
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal(expected, s.Num(formula), 9);
    }

    [Fact]
    public void Indirect_BadRef_IsRefError()
    {
        var s = Grid(out var wb);
        using (wb) Assert.Equal("#REF!", s.Str("INDIRECT(\"not a ref\")"));
    }
}
