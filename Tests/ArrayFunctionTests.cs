using TVE.PureDocs.Excel;
using Xunit;

namespace PureDocs.Excel.Tests;

public class ArrayFunctionTests
{
    private static Worksheet Data(out Workbook wb)
    {
        wb = Workbook.Create();
        var s = wb.AddWorksheet("S");
        // A1:A4 = 3,1,2,1 ; B1:B4 = x,y,x,z
        double[] a = { 3, 1, 2, 1 };
        string[] b = { "x", "y", "x", "z" };
        for (int i = 0; i < 4; i++) { s[$"A{i + 1}"].SetValue(a[i]); s[$"B{i + 1}"].SetValue(b[i]); }
        return s;
    }

    [Theory]
    [InlineData("SUM(SEQUENCE(3))", 6)]
    [InlineData("SUM(SEQUENCE(2,3))", 21)]
    [InlineData("SUM(SEQUENCE(3,1,10,5))", 45)]        // 10+15+20
    [InlineData("SUM(UNIQUE(A1:A4))", 6)]              // 3+1+2
    [InlineData("COUNTA(UNIQUE(B1:B4))", 3)]           // x,y,z
    [InlineData("SUM(FILTER(A1:A4,A1:A4>1))", 5)]      // 3+2
    [InlineData("INDEX(SORT(A1:A4),1,1)", 1)]          // ascending min
    [InlineData("INDEX(SORT(A1:A4,1,-1),1,1)", 3)]     // descending max
    [InlineData("COUNTA(TEXTSPLIT(\"a,b,c\",\",\"))", 3)]
    public void ArrayFunctionsAsValues(string formula, double expected)
    {
        var s = Data(out var wb);
        using (wb) Assert.Equal(expected, s.Num(formula), 9);
    }

    [Fact]
    public void Transpose_SumUnchanged()
    {
        var s = Data(out var wb);
        using (wb)
        {
            for (int r = 1; r <= 3; r++) s[$"D{r}"].SetValue(r * 1.0);
            Assert.Equal(6, s.Num("SUM(TRANSPOSE(D1:D3))"), 9);
        }
    }

    [Fact]
    public void MMult_MatrixTimesVector()
    {
        var s = Data(out var wb);
        using (wb)
        {
            s.GetRange("F1:G2").SetValues(new object[,] { { 1.0, 2.0 }, { 3.0, 4.0 } });
            s["H1"].SetValue(1.0); s["H2"].SetValue(1.0);
            // [1,2;3,4] x [1;1] = [3;7] -> sum 10
            Assert.Equal(10, s.Num("SUM(MMULT(F1:G2,H1:H2))"), 9);
        }
    }
}
