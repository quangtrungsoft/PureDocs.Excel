namespace TVE.PureDocs.Excel.Formulas;

internal static class InfoFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("ISBLANK", IsBlank, 1, 1); r.Register("ISNUMBER", IsNumber, 1, 1);
        r.Register("ISTEXT", IsText, 1, 1); r.Register("ISERROR", IsError, 1, 1);
        r.Register("ISERR", IsErr, 1, 1); r.Register("ISLOGICAL", IsLogical, 1, 1);
        r.Register("ISNA", IsNa, 1, 1); r.Register("NA", Na, 0, 0);
        r.Register("TYPE", TypeFn, 1, 1); r.Register("N", N, 1, 1);
        r.Register("ISODD", IsOdd, 1, 1); r.Register("ISEVEN", IsEven, 1, 1);
        r.Register("ISNONTEXT", IsNonText, 1, 1); r.Register("ERROR.TYPE", ErrorType, 1, 1);
    }

    private static FormulaValue IsBlank(List<FormulaNode> a, FormulaContext c)
        => FormulaValue.Boolean(a[0].Evaluate(c).IsBlank);

    private static FormulaValue IsNumber(List<FormulaNode> a, FormulaContext c)
        => FormulaValue.Boolean(a[0].Evaluate(c).IsNumber);

    private static FormulaValue IsText(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        return FormulaValue.Boolean(v.IsText && v.TextValue.Length > 0);
    }

    private static FormulaValue IsError(List<FormulaNode> a, FormulaContext c)
        => FormulaValue.Boolean(a[0].Evaluate(c).IsError);

    private static FormulaValue IsErr(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        return FormulaValue.Boolean(v.IsError && v.ErrorCode != FormulaError.NA);
    }

    private static FormulaValue IsLogical(List<FormulaNode> a, FormulaContext c)
        => FormulaValue.Boolean(a[0].Evaluate(c).IsBoolean);

    private static FormulaValue IsNa(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        return FormulaValue.Boolean(v.IsError && v.ErrorCode == FormulaError.NA);
    }

    private static FormulaValue Na(List<FormulaNode> a, FormulaContext c)
        => FormulaValue.ErrorNA;

    private static FormulaValue TypeFn(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        return FormulaValue.Number(v.Kind switch
        {
            FormulaValueKind.Number => 1, FormulaValueKind.Text => 2,
            FormulaValueKind.Boolean => 4, FormulaValueKind.Error => 16,
            FormulaValueKind.Array => 64, FormulaValueKind.Blank => 1, _ => 1,
        });
    }

    private static FormulaValue N(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        if (v.IsNumber) return v;
        if (v.IsBoolean) return FormulaValue.Number(v.BooleanValue ? 1 : 0);
        return FormulaValue.Zero;
    }

    private static FormulaValue IsOdd(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return FormulaValue.Boolean(Math.Abs((int)Math.Truncate(v)) % 2 == 1);
    }

    private static FormulaValue IsEven(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double v, out var e)) return e;
        return FormulaValue.Boolean((int)Math.Truncate(v) % 2 == 0);
    }

    private static FormulaValue IsNonText(List<FormulaNode> a, FormulaContext c)
        => FormulaValue.Boolean(!a[0].Evaluate(c).IsText);

    private static FormulaValue ErrorType(List<FormulaNode> a, FormulaContext c)
    {
        var v = a[0].Evaluate(c);
        if (!v.IsError) return FormulaValue.ErrorNA;
        return FormulaValue.Number(v.ErrorCode switch
        {
            FormulaError.Null => 1, FormulaError.Div0 => 2, FormulaError.Value => 3,
            FormulaError.Ref => 4, FormulaError.Name => 5, FormulaError.Num => 6,
            FormulaError.NA => 7, _ => 7,
        });
    }
}
