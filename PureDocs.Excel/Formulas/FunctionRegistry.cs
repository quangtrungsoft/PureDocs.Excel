namespace TVE.PureDocs.Excel.Formulas;

/// <summary>Function delegate returning FormulaValue (no boxing).</summary>
internal delegate FormulaValue FormulaFunction(List<FormulaNode> args, FormulaContext context);

/// <summary>Function metadata: arity, volatility, range support.</summary>
internal sealed class FunctionDefinition
{
    public required string Name { get; init; }
    public required int MinArgs { get; init; }
    public int MaxArgs { get; init; } = -1; // -1 = unlimited
    public bool IsVolatile { get; init; }
    /// <summary>If true, function can accept range references directly (e.g., SUM, AVERAGE).</summary>
    public bool AllowRange { get; init; }
    /// <summary>Optional description for IDE/autocomplete support.</summary>
    public string? Description { get; init; }
    public required FormulaFunction Handler { get; init; }
}

/// <summary>
/// Central registry of all supported Excel functions with metadata.
/// </summary>
internal sealed class FunctionRegistry
{
    private static FunctionRegistry? _default;
    private readonly Dictionary<string, FunctionDefinition> _functions;

    public static FunctionRegistry Default => _default ??= CreateDefault();

    private FunctionRegistry()
    {
        _functions = new Dictionary<string, FunctionDefinition>(StringComparer.OrdinalIgnoreCase);
    }

    /// <summary>Register a function with full metadata.</summary>
    public void Register(FunctionDefinition def)
    {
        _functions[def.Name.ToUpperInvariant()] = def;
    }

    /// <summary>Simple registration (backward compat helper).</summary>
    public void Register(string name, FormulaFunction handler, int minArgs = 0, int maxArgs = -1, bool isVolatile = false)
    {
        _functions[name.ToUpperInvariant()] = new FunctionDefinition
        {
            Name = name.ToUpperInvariant(),
            MinArgs = minArgs,
            MaxArgs = maxArgs,
            IsVolatile = isVolatile,
            Handler = handler
        };
    }

    /// <summary>Execute a function with arg validation.</summary>
    public FormulaValue Execute(string name, List<FormulaNode> args, FormulaContext context)
    {
        if (!_functions.TryGetValue(name, out var def))
            return FormulaValue.ErrorName;

        if (args.Count < def.MinArgs)
            return FormulaValue.ErrorValue;
        if (def.MaxArgs >= 0 && args.Count > def.MaxArgs)
            return FormulaValue.ErrorValue;

        return def.Handler(args, context);
    }

    /// <summary>Check if a function is volatile.</summary>
    public bool IsVolatile(string name)
    {
        return _functions.TryGetValue(name, out var def) && def.IsVolatile;
    }

    private static FunctionRegistry CreateDefault()
    {
        var registry = new FunctionRegistry();
        MathFunctions.Register(registry);
        TextFunctions.Register(registry);
        LogicalFunctions.Register(registry);
        DateFunctions.Register(registry);
        LookupFunctions.Register(registry);
        StatisticalFunctions.Register(registry);
        InfoFunctions.Register(registry);
        return registry;
    }
}
