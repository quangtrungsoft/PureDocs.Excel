namespace TVE.PureDocs.Excel.Formulas;

internal static class LogicalFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("IF", If, 2, 3); r.Register("AND", And, 1);
        r.Register("OR", Or, 1); r.Register("NOT", Not, 1, 1);
        r.Register("XOR", Xor, 1); r.Register("IFERROR", IfError, 2, 2);
        r.Register("IFNA", IfNa, 2, 2); r.Register("IFS", Ifs, 2);
        r.Register("SWITCH", Switch, 3); r.Register("TRUE", TrueFn, 0, 0);
        r.Register("FALSE", FalseFn, 0, 0);
    }

    private static FormulaValue If(List<FormulaNode> a, FormulaContext c)
    {
        var cond = a[0].Evaluate(c);
        if (cond.IsError) return cond;
        var b = cond.CoerceToBool();
        if (b.IsError) return b;
        return b.BooleanValue ? a[1].Evaluate(c) : a.Count > 2 ? a[2].Evaluate(c) : FormulaValue.False;
    }

    private static FormulaValue And(List<FormulaNode> a, FormulaContext c)
    {
        bool hasValue = false; // Track if we found any non-blank values
        foreach (var arg in a)
        {
            var v = arg.Evaluate(c);
            if (v.IsError) return v;
            if (v.IsArray)
            {
                foreach (var item in v.ArrayVal.EnumerateValues())
                {
                    // Excel's AND() ignores blank cells in ranges
                    if (item.IsBlank) continue;
                    if (item.IsError) return item;
                    hasValue = true;
                    var bv = item.CoerceToBool();
                    if (bv.IsError) return bv;
                    if (!bv.BooleanValue) return FormulaValue.False;
                }
            }
            else
            {
                // Single blank value is treated as FALSE for direct arguments
                if (v.IsBlank) 
                {
                    hasValue = true;
                    continue; // Blank direct arg is ignored, same as TRUE
                }
                hasValue = true;
                var bv = v.CoerceToBool();
                if (bv.IsError) return bv;
                if (!bv.BooleanValue) return FormulaValue.False;
            }
        }
        // If all values were blank, return #VALUE! (Excel behavior)
        return hasValue ? FormulaValue.True : FormulaValue.ErrorValue;
    }

    private static FormulaValue Or(List<FormulaNode> a, FormulaContext c)
    {
        bool hasValue = false; // Track if we found any non-blank values
        foreach (var arg in a)
        {
            var v = arg.Evaluate(c);
            if (v.IsError) return v;
            if (v.IsArray)
            {
                foreach (var item in v.ArrayVal.EnumerateValues())
                {
                    // Excel's OR() ignores blank cells in ranges
                    if (item.IsBlank) continue;
                    if (item.IsError) return item;
                    hasValue = true;
                    var bv = item.CoerceToBool();
                    if (bv.IsError) return bv;
                    if (bv.BooleanValue) return FormulaValue.True;
                }
            }
            else
            {
                // Single blank value is treated as FALSE for direct arguments
                if (v.IsBlank)
                {
                    hasValue = true;
                    continue; // Blank direct arg is ignored, same as FALSE
                }
                hasValue = true;
                var bv = v.CoerceToBool();
                if (bv.IsError) return bv;
                if (bv.BooleanValue) return FormulaValue.True;
            }
        }
        // If all values were blank, return #VALUE! (Excel behavior)
        return hasValue ? FormulaValue.False : FormulaValue.ErrorValue;
    }

    private static FormulaValue Not(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        if (v.IsError) return v;
        var bv = v.CoerceToBool();
        if (bv.IsError) return bv;
        return FormulaValue.Boolean(!bv.BooleanValue);
    }

    private static FormulaValue Xor(List<FormulaNode> a, FormulaContext c)
    {
        int trueCount = 0;
        foreach (var arg in a)
        {
            var v = arg.Evaluate(c);
            if (v.IsError) return v;
            var bv = v.CoerceToBool();
            if (bv.IsError) return bv;
            if (bv.BooleanValue) trueCount++;
        }
        return FormulaValue.Boolean(trueCount % 2 == 1);
    }

    private static FormulaValue IfError(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        return v.IsError ? a[1].Evaluate(c) : v;
    }

    private static FormulaValue IfNa(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        return v.IsError && v.ErrorCode == FormulaError.NA ? a[1].Evaluate(c) : v;
    }

    private static FormulaValue Ifs(List<FormulaNode> a, FormulaContext c)
    {
        for (int i = 0; i + 1 < a.Count; i += 2)
        {
            var cond = a[i].Evaluate(c);
            if (cond.IsError) return cond;
            var bv = cond.CoerceToBool();
            if (bv.IsError) return bv;
            if (bv.BooleanValue) return a[i + 1].Evaluate(c);
        }
        return FormulaValue.ErrorNA;
    }

    private static FormulaValue Switch(List<FormulaNode> a, FormulaContext c)
    {
        var expr = a[0].Evaluate(c);
        if (expr.IsError) return expr;
        for (int i = 1; i + 1 < a.Count; i += 2)
        {
            var cv = a[i].Evaluate(c);
            if (cv.IsError) return cv;
            if (FormulaValue.AreEqual(expr, cv)) return a[i + 1].Evaluate(c);
        }
        return a.Count % 2 == 0 ? a[^1].Evaluate(c) : FormulaValue.ErrorNA;
    }

    private static FormulaValue TrueFn(List<FormulaNode> a, FormulaContext c) => FormulaValue.True;
    private static FormulaValue FalseFn(List<FormulaNode> a, FormulaContext c) => FormulaValue.False;
}
