namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Main entry point for formula evaluation.
/// Includes AST cache to avoid re-parsing identical formulas.
/// </summary>
public static class FormulaEvaluator
{
    // AST cache: formula text → parsed AST (thread-safe)
    // Using LRU cache with bounded size to prevent memory leaks in long-running apps
    private static readonly LruCache<string, FormulaNode> _astCache = new(DefaultCacheCapacity);
    private const int DefaultCacheCapacity = 10_000;

    /// <summary>
    /// Gets or sets the maximum AST cache capacity.
    /// Setting this will clear the existing cache and create a new one with the specified capacity.
    /// Default is 10,000 entries.
    /// </summary>
    public static int MaxCacheCapacity
    {
        get => _maxCacheCapacity;
        set
        {
            if (value < 100) value = 100; // Minimum reasonable size
            if (value != _maxCacheCapacity)
            {
                _maxCacheCapacity = value;
                _astCache.Resize(value);
            }
        }
    }
    private static int _maxCacheCapacity = DefaultCacheCapacity;

    /// <summary>Evaluates a formula string against a worksheet.</summary>
    public static FormulaValue Evaluate(string formula, Worksheet worksheet,
        NamedRangeManager? namedRanges = null)
    {
        if (string.IsNullOrWhiteSpace(formula))
            return FormulaValue.Blank;

        formula = formula.TrimStart('=');

        try
        {
            var ast = GetOrParseAst(formula);
            var context = new FormulaContext(worksheet, namedRanges: namedRanges);
            return ast.Evaluate(context);
        }
        catch (FormulaException ex)
        {
            return FormulaValue.Error(FormulaValue.ErrorFromString(ex.Message));
        }
        catch (Exception)
        {
            return FormulaValue.ErrorValue;
        }
    }

    /// <summary>Internal evaluate with shared evaluatingCells set (for recursive cell evaluation).</summary>
    internal static FormulaValue EvaluateInternal(string formula, Worksheet worksheet,
        HashSet<string> evaluatingCells, NamedRangeManager? namedRanges = null)
    {
        if (string.IsNullOrWhiteSpace(formula))
            return FormulaValue.Blank;

        formula = formula.TrimStart('=');

        try
        {
            var ast = GetOrParseAst(formula);
            var context = new FormulaContext(worksheet, evaluatingCells, namedRanges);
            return ast.Evaluate(context);
        }
        catch (FormulaException ex)
        {
            return FormulaValue.Error(FormulaValue.ErrorFromString(ex.Message));
        }
        catch (Exception)
        {
            return FormulaValue.ErrorValue;
        }
    }

    /// <summary>Parses formula to AST, with caching.</summary>
    internal static FormulaNode GetOrParseAst(string formula)
    {
        // Try to get from cache first
        if (_astCache.TryGet(formula, out var cached))
            return cached!;

        // Parse and cache
        var lexer = new FormulaLexer(formula);
        var tokens = lexer.Tokenize();
        var ast = new FormulaParser(tokens).Parse();
        
        _astCache.Set(formula, ast);
        return ast;
    }

    /// <summary>Clears the AST cache (useful after bulk edits or to free memory).</summary>
    public static void ClearCache() => _astCache.Clear();

    /// <summary>Current cache size (for diagnostics).</summary>
    public static int CacheSize => _astCache.Count;

    /// <summary>Gets cache statistics: hits, misses, evictions, hit rate.</summary>
    public static (long Hits, long Misses, long Evictions, double HitRate) GetCacheStats()
        => (_astCache.Hits, _astCache.Misses, _astCache.Evictions, _astCache.HitRate);
}
