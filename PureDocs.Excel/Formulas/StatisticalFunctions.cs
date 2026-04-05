namespace TVE.PureDocs.Excel.Formulas;

internal static class StatisticalFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("MEDIAN", Median, 1); r.Register("MODE", Mode, 1);
        r.Register("STDEV", Stdev, 1); r.Register("STDEV.S", Stdev, 1);
        r.Register("STDEVP", Stdevp, 1); r.Register("STDEV.P", Stdevp, 1);
        r.Register("VAR", Var, 1); r.Register("VAR.S", Var, 1);
        r.Register("VARP", Varp, 1); r.Register("VAR.P", Varp, 1);
        r.Register("LARGE", Large, 2, 2); r.Register("SMALL", Small, 2, 2);
        r.Register("RANK", Rank, 2, 3); r.Register("RANK.EQ", Rank, 2, 3);
        r.Register("PERCENTILE", Percentile, 2, 2); r.Register("PERCENTILE.INC", Percentile, 2, 2);
        r.Register("MAXIFS", MaxIfs, 3); r.Register("MINIFS", MinIfs, 3);
    }

    private static FormulaValue Median(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var e)) return e;
        if (nums.Count == 0) return FormulaValue.ErrorNum;
        nums.Sort();
        int mid = nums.Count / 2;
        return FormulaValue.Number(nums.Count % 2 == 0 ? (nums[mid - 1] + nums[mid]) / 2 : nums[mid]);
    }

    /// <summary>
    /// MODE - Returns the most frequently occurring value.
    /// For multimodal datasets, returns the first mode in order of appearance (Excel behavior).
    /// </summary>
    private static FormulaValue Mode(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var e)) return e;
        if (nums.Count == 0) return FormulaValue.ErrorNA;
        
        var freq = new Dictionary<double, int>();
        foreach (var n in nums) { freq.TryGetValue(n, out int cnt); freq[n] = cnt + 1; }
        int maxCnt = freq.Values.Max();
        if (maxCnt == 1) return FormulaValue.ErrorNA; // No repeating values
        
        // Excel returns first occurrence of the mode in data order
        // Iterate through original data to find first value with max frequency
        foreach (var n in nums)
        {
            if (freq[n] == maxCnt)
                return FormulaValue.Number(n);
        }
        
        return FormulaValue.ErrorNA; // Should not reach here
    }

    private static FormulaValue Stdev(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var e)) return e;
        if (nums.Count < 2) return FormulaValue.ErrorDiv0;
        return FormulaValue.Number(Math.Sqrt(Variance(nums, false)));
    }

    private static FormulaValue Stdevp(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var e)) return e;
        if (nums.Count == 0) return FormulaValue.ErrorDiv0;
        return FormulaValue.Number(Math.Sqrt(Variance(nums, true)));
    }

    private static FormulaValue Var(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var e)) return e;
        if (nums.Count < 2) return FormulaValue.ErrorDiv0;
        return FormulaValue.Number(Variance(nums, false));
    }

    private static FormulaValue Varp(List<FormulaNode> a, FormulaContext c)
    {
        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(a, c, nums, out var e)) return e;
        if (nums.Count == 0) return FormulaValue.ErrorDiv0;
        return FormulaValue.Number(Variance(nums, true));
    }

    private static double Variance(List<double> nums, bool population)
    {
        double mean = nums.Average();
        double sumSq = 0;
        foreach (var n in nums) sumSq += (n - mean) * (n - mean);
        return sumSq / (population ? nums.Count : nums.Count - 1);
    }

    private static FormulaValue Large(List<FormulaNode> a, FormulaContext c)
    {
        var arrVal = a[0].Evaluate(c);
        if (arrVal.IsError) return arrVal;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double kv, out var e)) return e;
        int k = (int)kv;
        var nums = CollectFromArray(arrVal);
        if (k < 1 || k > nums.Count) return FormulaValue.ErrorNum;
        nums.Sort(); nums.Reverse();
        return FormulaValue.Number(nums[k - 1]);
    }

    private static FormulaValue Small(List<FormulaNode> a, FormulaContext c)
    {
        var arrVal = a[0].Evaluate(c);
        if (arrVal.IsError) return arrVal;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double kv, out var e)) return e;
        int k = (int)kv;
        var nums = CollectFromArray(arrVal);
        if (k < 1 || k > nums.Count) return FormulaValue.ErrorNum;
        nums.Sort();
        return FormulaValue.Number(nums[k - 1]);
    }

    private static FormulaValue Rank(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double num, out var e)) return e;
        var arrVal = a[1].Evaluate(c);
        if (arrVal.IsError) return arrVal;
        int order = 0;
        if (a.Count > 2) { if (!FormulaHelper.TryEvalDouble(a[2], c, out double ov, out e)) return e; order = (int)ov; }
        var nums = CollectFromArray(arrVal);
        nums.Sort(); if (order == 0) nums.Reverse();
        int rank = nums.IndexOf(num);
        return rank < 0 ? FormulaValue.ErrorNA : FormulaValue.Number(rank + 1);
    }

    private static FormulaValue Percentile(List<FormulaNode> a, FormulaContext c)
    {
        var arrVal = a[0].Evaluate(c);
        if (arrVal.IsError) return arrVal;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double k, out var e)) return e;
        if (k < 0 || k > 1) return FormulaValue.ErrorNum;
        var nums = CollectFromArray(arrVal);
        if (nums.Count == 0) return FormulaValue.ErrorNum;
        nums.Sort();
        double n = (nums.Count - 1) * k;
        int lo = (int)Math.Floor(n), hi = (int)Math.Ceiling(n);
        if (lo == hi) return FormulaValue.Number(nums[lo]);
        return FormulaValue.Number(nums[lo] + (nums[hi] - nums[lo]) * (n - lo));
    }

    private static FormulaValue MaxIfs(List<FormulaNode> a, FormulaContext c)
    {
        if (a.Count < 3 || (a.Count - 1) % 2 != 0) return FormulaValue.ErrorValue;
        var maxRange = a[0].Evaluate(c);
        if (maxRange.IsError) return maxRange;
        if (!maxRange.IsArray) return FormulaValue.ErrorValue;
        var maxArr = maxRange.ArrayVal;
        int len = maxArr.Length;

        bool[] mask = new bool[len];
        Array.Fill(mask, true);

        for (int p = 1; p + 1 < a.Count; p += 2)
        {
            var cr = a[p].Evaluate(c);
            if (cr.IsError) return cr;
            var criteria = a[p + 1].Evaluate(c);
            if (criteria.IsError) return criteria;
            string crit = criteria.AsText();
            var crArr = cr.IsArray ? cr.ArrayVal : null;
            for (int i = 0; i < len; i++)
            {
                if (!mask[i]) continue;
                var cv = crArr != null ? crArr[i] : cr;
                if (!FormulaHelper.MatchesCriteria(cv, crit)) mask[i] = false;
            }
        }

        double max = double.NegativeInfinity; bool found = false;
        for (int i = 0; i < len; i++)
            if (mask[i] && maxArr[i].TryAsDouble(out double d)) { if (d > max) max = d; found = true; }
        return found ? FormulaValue.Number(max) : FormulaValue.Zero;
    }

    private static FormulaValue MinIfs(List<FormulaNode> a, FormulaContext c)
    {
        if (a.Count < 3 || (a.Count - 1) % 2 != 0) return FormulaValue.ErrorValue;
        var minRange = a[0].Evaluate(c);
        if (minRange.IsError) return minRange;
        if (!minRange.IsArray) return FormulaValue.ErrorValue;
        var minArr = minRange.ArrayVal;
        int len = minArr.Length;

        bool[] mask = new bool[len];
        Array.Fill(mask, true);

        for (int p = 1; p + 1 < a.Count; p += 2)
        {
            var cr = a[p].Evaluate(c);
            if (cr.IsError) return cr;
            var criteria = a[p + 1].Evaluate(c);
            if (criteria.IsError) return criteria;
            string crit = criteria.AsText();
            var crArr = cr.IsArray ? cr.ArrayVal : null;
            for (int i = 0; i < len; i++)
            {
                if (!mask[i]) continue;
                var cv = crArr != null ? crArr[i] : cr;
                if (!FormulaHelper.MatchesCriteria(cv, crit)) mask[i] = false;
            }
        }

        double min = double.PositiveInfinity; bool found = false;
        for (int i = 0; i < len; i++)
            if (mask[i] && minArr[i].TryAsDouble(out double d)) { if (d < min) min = d; found = true; }
        return found ? FormulaValue.Number(min) : FormulaValue.Zero;
    }

    private static List<double> CollectFromArray(FormulaValue v)
    {
        var nums = new List<double>();
        if (v.IsArray)
            foreach (var item in v.ArrayVal.EnumerateValues())
            { if (item.IsNumber) nums.Add(item.NumberValue); }
        else if (v.IsNumber) nums.Add(v.NumberValue);
        return nums;
    }
}
