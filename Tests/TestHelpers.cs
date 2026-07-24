using TVE.PureDocs.Excel;
using TVE.PureDocs.Excel.Formulas;

namespace PureDocs.Excel.Tests;

/// <summary>Shared helpers for the formula/workbook test suite.</summary>
internal static class TestHelpers
{
    /// <summary>Evaluates a formula on a worksheet and coerces the result to double.</summary>
    public static double Num(this Worksheet sheet, string formula)
        => Convert.ToDouble(sheet.EvaluateFormula(formula));

    /// <summary>Evaluates a formula and returns its text form (e.g. "#SPILL!", "hello").</summary>
    public static string Str(this Worksheet sheet, string formula)
        => sheet.EvaluateFormula(formula)?.ToString() ?? "";

    /// <summary>The cached recalc value for a cell as a double (NaN if uncached).</summary>
    public static double Cached(this Worksheet sheet, string reference)
    {
        var v = sheet.CalcChain.GetCachedValue(CellAddress.FromReference(reference));
        return v.HasValue ? v.Value.NumberValue : double.NaN;
    }

    /// <summary>The spilled value covering a cell as a double, or null if not spilled.</summary>
    public static double? Spill(this Worksheet sheet, string reference)
    {
        var v = sheet.CalcChain.Spills.GetSpillValue(CellAddress.FromReference(reference));
        return v.HasValue ? v.Value.NumberValue : (double?)null;
    }

    /// <summary>Creates a unique temp .xlsx path (caller deletes).</summary>
    public static string TempXlsx()
        => Path.Combine(Path.GetTempPath(), $"pd_test_{Guid.NewGuid():N}.xlsx");
}
