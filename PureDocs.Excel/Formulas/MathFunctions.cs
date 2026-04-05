namespace TVE.PureDocs.Excel.Formulas;

internal static class MathFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("SUM", Sum, 1); r.Register("AVERAGE", Average, 1);
        r.Register("MIN", Min, 1); r.Register("MAX", Max, 1);
        r.Register("COUNT", Count, 1); r.Register("COUNTA", CountA, 1);
        r.Register("COUNTBLANK", CountBlank, 1, 1);
        r.Register("ABS", Abs, 1, 1); r.Register("ROUND", Round, 2, 2);
        r.Register("ROUNDUP", RoundUp, 2, 2); r.Register("ROUNDDOWN", RoundDown, 2, 2);
        r.Register("CEILING", Ceiling, 2, 2); r.Register("FLOOR", Floor, 2, 2);
        r.Register("INT", Int, 1, 1); r.Register("MOD", Mod, 2, 2);
        r.Register("POWER", Power, 2, 2); r.Register("SQRT", Sqrt, 1, 1);
        r.Register("SIGN", Sign, 1, 1); r.Register("PI", Pi, 0, 0);
        r.Register("RAND", Rand, 0, 0, isVolatile: true);
        r.Register("RANDBETWEEN", RandBetween, 2, 2, isVolatile: true);
        r.Register("LOG", Log, 1, 2); r.Register("LOG10", Log10, 1, 1);
        r.Register("LN", Ln, 1, 1); r.Register("EXP", Exp, 1, 1);
        r.Register("FACT", Fact, 1, 1); r.Register("COMBIN", Combin, 2, 2);
        r.Register("PRODUCT", Product, 1); r.Register("SUMPRODUCT", SumProduct, 1);
        r.Register("SUMIF", SumIf, 2, 3); r.Register("COUNTIF", CountIf, 2, 2);
        r.Register("AVERAGEIF", AverageIf, 2, 3);
        r.Register("TRUNC", Trunc, 1, 2); r.Register("QUOTIENT", Quotient, 2, 2);
    }

    // ── Aggregates ──────────────────────────────────────────────────

    private static FormulaValue Sum(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var err)) return err;
        double s = 0; foreach (var n in nums) s += n;
        return FormulaValue.Number(s);
    }

    private static FormulaValue Average(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var err)) return err;
        return nums.Count == 0 ? FormulaValue.ErrorDiv0 : FormulaValue.Number(nums.Sum() / nums.Count);
    }

    private static FormulaValue Min(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var err)) return err;
        return nums.Count == 0 ? FormulaValue.Zero : FormulaValue.Number(nums.Min());
    }

    private static FormulaValue Max(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var err)) return err;
        return nums.Count == 0 ? FormulaValue.Zero : FormulaValue.Number(nums.Max());
    }

    private static FormulaValue Count(List<FormulaNode> a, FormulaContext c)
    {
        int count = 0;
        var vals = new List<FormulaValue>();
        if (!FormulaHelper.TryFlattenArgs(a, c, vals, out _))
        {
            // Count doesn't propagate errors — it counts
        }
        else
            foreach (var v in vals) { if (v.IsNumber || (v.IsText && v.TryAsDouble(out _))) count++; }
        return FormulaValue.Number(count);
    }

    private static FormulaValue CountA(List<FormulaNode> a, FormulaContext c)
    {
        var vals = new List<FormulaValue>();
        FormulaHelper.TryFlattenArgs(a, c, vals, out _);
        int count = 0;
        foreach (var v in vals) { if (!v.IsBlank) count++; }
        return FormulaValue.Number(count);
    }

    private static FormulaValue CountBlank(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        if (!v.IsArray) return v.IsBlank ? FormulaValue.One : FormulaValue.Zero;
        int count = 0;
        foreach (var item in v.ArrayVal.EnumerateValues()) { if (item.IsBlank) count++; }
        return FormulaValue.Number(count);
    }

    private static FormulaValue Product(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var err)) return err;
        double p = 1; foreach (var n in nums) p *= n;
        return FormulaValue.Number(p);
    }

    // ── Single-arg math ─────────────────────────────────────────────

    private static FormulaValue Abs(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return FormulaValue.Number(Math.Abs(v));
    }

    private static FormulaValue Sqrt(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return v < 0 ? FormulaValue.ErrorNum : FormulaValue.Number(Math.Sqrt(v));
    }

    private static FormulaValue Sign(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return FormulaValue.Number(Math.Sign(v));
    }

    private static FormulaValue Int(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return FormulaValue.Number(Math.Floor(v));
    }

    private static FormulaValue Ln(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return v <= 0 ? FormulaValue.ErrorNum : FormulaValue.Number(Math.Log(v));
    }

    private static FormulaValue Log10(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return v <= 0 ? FormulaValue.ErrorNum : FormulaValue.Number(Math.Log10(v));
    }

    private static FormulaValue Exp(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return FormulaValue.Number(Math.Exp(v));
    }

    private static FormulaValue Fact(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        int n = (int)Math.Floor(v);
        if (n < 0) return FormulaValue.ErrorNum;
        double r = 1; for (int i = 2; i <= n; i++) r *= i;
        return FormulaValue.Number(r);
    }

    private static FormulaValue Pi(List<FormulaNode> a, FormulaContext c) => FormulaValue.Number(Math.PI);

    // ── Two-arg math ────────────────────────────────────────────────

    private static FormulaValue Round(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double d, out e)) return e;
        return FormulaValue.Number(Math.Round(v, Math.Max(0, (int)d), MidpointRounding.AwayFromZero));
    }

    private static FormulaValue RoundUp(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double d, out e)) return e;
        double f = Math.Pow(10, (int)d);
        return FormulaValue.Number(Math.Ceiling(Math.Abs(v) * f) / f * Math.Sign(v));
    }

    private static FormulaValue RoundDown(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double d, out e)) return e;
        double f = Math.Pow(10, (int)d);
        return FormulaValue.Number(Math.Truncate(v * f) / f);
    }

    private static FormulaValue Ceiling(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double s, out e)) return e;
        if (s == 0) return FormulaValue.Zero;
        return FormulaValue.Number(Math.Ceiling(v / s) * s);
    }

    private static FormulaValue Floor(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double s, out e)) return e;
        if (s == 0) return FormulaValue.ErrorDiv0;
        return FormulaValue.Number(Math.Floor(v / s) * s);
    }

    private static FormulaValue Mod(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double d, out e)) return e;
        if (d == 0) return FormulaValue.ErrorDiv0;
        return FormulaValue.Number(v - d * Math.Floor(v / d));
    }

    private static FormulaValue Power(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double b, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double p, out e)) return e;
        return FormulaValue.Number(Math.Pow(b, p));
    }

    private static FormulaValue Log(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        if (v <= 0) return FormulaValue.ErrorNum;
        double b = 10; if (a.Count > 1) { if (!FormulaHelper.TryEvalDouble(a[1], c, out b, out e)) return e; }
        if (b <= 0 || b == 1) return FormulaValue.ErrorNum;
        return FormulaValue.Number(Math.Log(v, b));
    }

    private static FormulaValue Combin(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double nv, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double kv, out e)) return e;
        int n = (int)nv, k = (int)kv;
        if (n < 0 || k < 0 || k > n) return FormulaValue.ErrorNum;
        double r = 1; for (int i = 0; i < k; i++) r = r * (n - i) / (i + 1);
        return FormulaValue.Number(Math.Round(r));
    }

    private static FormulaValue Trunc(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        int digits = 0; if (a.Count > 1) { if (!FormulaHelper.TryEvalDouble(a[1], c, out double d, out e)) return e; digits = (int)d; }
        double f = Math.Pow(10, digits);
        return FormulaValue.Number(Math.Truncate(v * f) / f);
    }

    private static FormulaValue Quotient(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double n, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double d, out e)) return e;
        if (d == 0) return FormulaValue.ErrorDiv0;
        return FormulaValue.Number(Math.Truncate(n / d));
    }

    // ── Random ──────────────────────────────────────────────────────

    private static readonly Random _rand = new();

    private static FormulaValue Rand(List<FormulaNode> a, FormulaContext c) => FormulaValue.Number(_rand.NextDouble());

    private static FormulaValue RandBetween(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double lo, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double hi, out e)) return e;
        int low = (int)Math.Ceiling(lo), high = (int)Math.Floor(hi);
        return low > high ? FormulaValue.ErrorNum : FormulaValue.Number(_rand.Next(low, high + 1));
    }

    // ── SUMPRODUCT ──────────────────────────────────────────────────

    private static FormulaValue SumProduct(List<FormulaNode> a, FormulaContext c)
    {
        var arrays = new List<ArrayValue>();
        foreach (var arg in a)
        {
            var v = arg.Evaluate(c);
            if (v.IsError) return v;
            if (!v.IsArray) { var sv = new ArrayValue(1, 1); sv[0] = v; arrays.Add(sv); }
            else arrays.Add(v.ArrayVal);
        }
        if (arrays.Count == 0) return FormulaValue.Zero;
        int len = arrays[0].Length;
        foreach (var arr in arrays) { if (arr.Length != len) return FormulaValue.ErrorValue; }
        double sum = 0;
        for (int i = 0; i < len; i++)
        {
            double prod = 1;
            foreach (var arr in arrays)
            {
                var item = arr[i];
                if (item.TryAsDouble(out double d)) prod *= d; else prod = 0;
            }
            sum += prod;
        }
        return FormulaValue.Number(sum);
    }

    // ── Conditional aggregates ──────────────────────────────────────

    private static FormulaValue SumIf(List<FormulaNode> a, FormulaContext c)
    {
        var rangeVal = a[0].Evaluate(c);
        if (rangeVal.IsError) return rangeVal;
        var criteriaVal = a[1].Evaluate(c);
        if (criteriaVal.IsError) return criteriaVal;
        string criteria = criteriaVal.AsText();
        var sumRangeVal = a.Count > 2 ? a[2].Evaluate(c) : rangeVal;
        if (sumRangeVal.IsError) return sumRangeVal;

        var range = rangeVal.IsArray ? rangeVal.ArrayVal : WrapSingle(rangeVal);
        var sumRange = sumRangeVal.IsArray ? sumRangeVal.ArrayVal : WrapSingle(sumRangeVal);
        double sum = 0;
        int len = Math.Min(range.Length, sumRange.Length);
        for (int i = 0; i < len; i++)
            if (FormulaHelper.MatchesCriteria(range[i], criteria) && sumRange[i].TryAsDouble(out double d)) sum += d;
        return FormulaValue.Number(sum);
    }

    private static FormulaValue CountIf(List<FormulaNode> a, FormulaContext c)
    {
        var rangeVal = a[0].Evaluate(c);
        if (rangeVal.IsError) return rangeVal;
        var criteriaVal = a[1].Evaluate(c);
        if (criteriaVal.IsError) return criteriaVal;
        string criteria = criteriaVal.AsText();
        var range = rangeVal.IsArray ? rangeVal.ArrayVal : WrapSingle(rangeVal);
        int count = 0;
        for (int i = 0; i < range.Length; i++)
            if (FormulaHelper.MatchesCriteria(range[i], criteria)) count++;
        return FormulaValue.Number(count);
    }

    private static FormulaValue AverageIf(List<FormulaNode> a, FormulaContext c)
    {
        var rangeVal = a[0].Evaluate(c);
        if (rangeVal.IsError) return rangeVal;
        var criteriaVal = a[1].Evaluate(c);
        if (criteriaVal.IsError) return criteriaVal;
        string criteria = criteriaVal.AsText();
        var avgRangeVal = a.Count > 2 ? a[2].Evaluate(c) : rangeVal;
        if (avgRangeVal.IsError) return avgRangeVal;

        var range = rangeVal.IsArray ? rangeVal.ArrayVal : WrapSingle(rangeVal);
        var avgRange = avgRangeVal.IsArray ? avgRangeVal.ArrayVal : WrapSingle(avgRangeVal);
        double sum = 0; int cnt = 0;
        int len = Math.Min(range.Length, avgRange.Length);
        for (int i = 0; i < len; i++)
            if (FormulaHelper.MatchesCriteria(range[i], criteria) && avgRange[i].TryAsDouble(out double d)) { sum += d; cnt++; }
        return cnt == 0 ? FormulaValue.ErrorDiv0 : FormulaValue.Number(sum / cnt);
    }

    private static ArrayValue WrapSingle(FormulaValue v)
    {
        var arr = new ArrayValue(1, 1); arr[0] = v; return arr;
    }
}
