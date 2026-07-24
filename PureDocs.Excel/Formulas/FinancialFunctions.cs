namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Time-value-of-money financial functions (PMT, FV, PV, NPV, IRR).
/// Sign and annuity conventions follow Excel: outflows are negative, and
/// <c>type</c> is 0 for end-of-period payments (default) or 1 for beginning.
/// </summary>
internal static class FinancialFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("PMT", Pmt, 3, 5);
        r.Register("FV", Fv, 3, 5);
        r.Register("PV", Pv, 3, 5);
        r.Register("NPV", Npv, 2);
        r.Register("IRR", Irr, 1, 2);
    }

    private static FormulaValue Pmt(List<FormulaNode> a, FormulaContext c)
    {
        if (!Read(a, c, out var rate, out var nper, out var pv, 0, 1, 2, out var e)) return e;
        double fv = 0, type = 0;
        if (a.Count > 3 && !FormulaHelper.TryEvalDouble(a[3], c, out fv, out e)) return e;
        if (a.Count > 4 && !FormulaHelper.TryEvalDouble(a[4], c, out type, out e)) return e;

        if (nper == 0) return FormulaValue.ErrorNum;
        if (rate == 0) return FormulaValue.Number(-(pv + fv) / nper);
        double pow = Math.Pow(1 + rate, nper);
        double pmt = -(pv * pow + fv) * rate / ((pow - 1) * (1 + rate * type));
        return FormulaValue.Number(pmt);
    }

    private static FormulaValue Fv(List<FormulaNode> a, FormulaContext c)
    {
        // FV(rate, nper, pmt, [pv], [type])
        if (!Read(a, c, out var rate, out var nper, out var pmt, 0, 1, 2, out var e)) return e;
        double pv = 0, type = 0;
        if (a.Count > 3 && !FormulaHelper.TryEvalDouble(a[3], c, out pv, out e)) return e;
        if (a.Count > 4 && !FormulaHelper.TryEvalDouble(a[4], c, out type, out e)) return e;

        if (rate == 0) return FormulaValue.Number(-(pv + pmt * nper));
        double pow = Math.Pow(1 + rate, nper);
        double fv = -(pv * pow + pmt * (1 + rate * type) * (pow - 1) / rate);
        return FormulaValue.Number(fv);
    }

    private static FormulaValue Pv(List<FormulaNode> a, FormulaContext c)
    {
        // PV(rate, nper, pmt, [fv], [type])
        if (!Read(a, c, out var rate, out var nper, out var pmt, 0, 1, 2, out var e)) return e;
        double fv = 0, type = 0;
        if (a.Count > 3 && !FormulaHelper.TryEvalDouble(a[3], c, out fv, out e)) return e;
        if (a.Count > 4 && !FormulaHelper.TryEvalDouble(a[4], c, out type, out e)) return e;

        if (rate == 0) return FormulaValue.Number(-(fv + pmt * nper));
        double pow = Math.Pow(1 + rate, nper);
        double pv = -(fv + pmt * (1 + rate * type) * (pow - 1) / rate) / pow;
        return FormulaValue.Number(pv);
    }

    private static FormulaValue Npv(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double rate, out var e)) return e;
        if (rate <= -1) return FormulaValue.ErrorNum;

        var flows = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a.GetRange(1, a.Count - 1), c, flows, out var err)) return err;

        double npv = 0;
        for (int i = 0; i < flows.Count; i++)
            npv += flows[i] / Math.Pow(1 + rate, i + 1);
        return FormulaValue.Number(npv);
    }

    private static FormulaValue Irr(List<FormulaNode> a, FormulaContext c)
    {
        var flows = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(new List<FormulaNode> { a[0] }, c, flows, out var err)) return err;
        if (flows.Count < 2) return FormulaValue.ErrorNum;

        double guess = 0.1;
        if (a.Count > 1 && !FormulaHelper.TryEvalDouble(a[1], c, out guess, out var e)) return e;

        // Newton-Raphson on NPV(rate) with periods starting at 0.
        double rate = guess;
        for (int iter = 0; iter < 100; iter++)
        {
            double npv = 0, dnpv = 0;
            for (int t = 0; t < flows.Count; t++)
            {
                double denom = Math.Pow(1 + rate, t);
                npv += flows[t] / denom;
                dnpv -= t * flows[t] / (denom * (1 + rate));
            }
            if (Math.Abs(npv) < 1e-7) return FormulaValue.Number(rate);
            if (dnpv == 0) break; // flat derivative — Newton cannot proceed
            double next = rate - npv / dnpv;
            if (double.IsNaN(next) || double.IsInfinity(next) || next <= -1) break;
            if (Math.Abs(next - rate) < 1e-9) { rate = next; break; }
            rate = next;
        }

        // Verify convergence before returning.
        double check = 0;
        for (int t = 0; t < flows.Count; t++) check += flows[t] / Math.Pow(1 + rate, t);
        return Math.Abs(check) < 1e-5 ? FormulaValue.Number(rate) : FormulaValue.ErrorNum;
    }

    /// <summary>Reads three required leading numeric arguments at the given indices.</summary>
    private static bool Read(List<FormulaNode> a, FormulaContext c,
        out double v0, out double v1, out double v2, int i0, int i1, int i2, out FormulaValue error)
    {
        v0 = v1 = v2 = 0;
        if (!FormulaHelper.TryEvalDouble(a[i0], c, out v0, out error)) return false;
        if (!FormulaHelper.TryEvalDouble(a[i1], c, out v1, out error)) return false;
        if (!FormulaHelper.TryEvalDouble(a[i2], c, out v2, out error)) return false;
        return true;
    }
}
