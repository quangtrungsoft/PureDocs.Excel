namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// A LAMBDA definition: parameter names, the body expression, and the local bindings
/// captured at creation time (so a lambda created inside LET closes over those names).
/// </summary>
internal sealed class LambdaValue
{
    public string[] Parameters { get; }
    public FormulaNode Body { get; }
    public Dictionary<string, FormulaValue> Captured { get; }

    public LambdaValue(string[] parameters, FormulaNode body, Dictionary<string, FormulaValue> captured)
    {
        Parameters = parameters;
        Body = body;
        Captured = captured;
    }
}

/// <summary>LET and LAMBDA — name binding and anonymous functions (Excel 365).</summary>
internal static class ScopeFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("LET", Let, 3);
        r.Register("LAMBDA", Lambda, 1);
    }

    /// <summary>LET(name1, value1, [name2, value2, ...], calculation).</summary>
    private static FormulaValue Let(List<FormulaNode> a, FormulaContext c)
    {
        // Need name/value pairs plus a final calculation → odd count, at least 3.
        if (a.Count < 3 || a.Count % 2 == 0) return FormulaValue.ErrorValue;

        var frame = new Dictionary<string, FormulaValue>(StringComparer.OrdinalIgnoreCase);
        c.PushScope(frame);
        try
        {
            for (int i = 0; i < a.Count - 1; i += 2)
            {
                if (a[i] is not NamedRangeNode nameNode) return FormulaValue.ErrorName;
                // Bind incrementally so later values can reference earlier names.
                var val = a[i + 1].Evaluate(c);
                if (val.IsError) return val;
                frame[nameNode.Name] = val;
            }
            return a[^1].Evaluate(c);
        }
        finally
        {
            c.PopScope();
        }
    }

    /// <summary>LAMBDA([param1, ...], calculation) — returns a callable lambda value.</summary>
    private static FormulaValue Lambda(List<FormulaNode> a, FormulaContext c)
    {
        int paramCount = a.Count - 1; // last arg is the body
        var names = new string[paramCount];
        for (int i = 0; i < paramCount; i++)
        {
            if (a[i] is not NamedRangeNode nameNode) return FormulaValue.ErrorName;
            names[i] = nameNode.Name;
        }
        // Capture current LET bindings so the lambda closes over them.
        return FormulaValue.Lambda(new LambdaValue(names, a[^1], c.SnapshotScope()));
    }
}
