using System.Buffers;

namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Pool for reusing FormulaValue arrays in function argument processing.
/// Avoids frequent allocation of temporary arrays during formula evaluation.
/// </summary>
public static class FormulaValuePool
{
    private static readonly ArrayPool<FormulaValue> _pool = ArrayPool<FormulaValue>.Create(
        maxArrayLength: 65536, maxArraysPerBucket: 32);

    private static readonly ArrayPool<double> _doublePool = ArrayPool<double>.Create(
        maxArrayLength: 65536, maxArraysPerBucket: 32);

    /// <summary>Rents a FormulaValue array from the pool.</summary>
    public static FormulaValue[] RentValues(int minimumLength) => _pool.Rent(minimumLength);

    /// <summary>Returns a FormulaValue array to the pool.</summary>
    public static void ReturnValues(FormulaValue[] array) => _pool.Return(array, clearArray: true);

    /// <summary>Rents a double array from the pool.</summary>
    public static double[] RentDoubles(int minimumLength) => _doublePool.Rent(minimumLength);

    /// <summary>Returns a double array to the pool.</summary>
    public static void ReturnDoubles(double[] array) => _doublePool.Return(array, clearArray: false);
}

/// <summary>
/// String interning for frequently used cell reference strings.
/// Reduces memory usage when the same cell references appear in many formulas.
/// </summary>
public static class CellReferenceIntern
{
    // Cache for commonly referenced cells (e.g., A1..Z100)
    private static readonly Dictionary<string, string> _cache = new(StringComparer.OrdinalIgnoreCase);
    private static readonly object _lock = new();

    /// <summary>Returns an interned version of the cell reference string.</summary>
    public static string Intern(string cellRef)
    {
        lock (_lock)
        {
            if (_cache.TryGetValue(cellRef, out var cached))
                return cached;

            // Only intern up to a reasonable limit to avoid unbounded memory growth
            if (_cache.Count < 100_000)
                _cache[cellRef] = cellRef;

            return cellRef;
        }
    }

    /// <summary>Clears the intern cache.</summary>
    public static void Clear()
    {
        lock (_lock) _cache.Clear();
    }

    /// <summary>Current cache size.</summary>
    public static int CacheSize
    {
        get { lock (_lock) return _cache.Count; }
    }
}

/// <summary>
/// Performance diagnostics for formula evaluation.
/// </summary>
public sealed class FormulaPerformanceCounters
{
    private long _parseCount;
    private long _evalCount;
    private long _cacheHits;
    private long _cacheMisses;
    private long _recalcTimeMs;

    public long ParseCount => _parseCount;
    public long EvalCount => _evalCount;
    public long CacheHits => _cacheHits;
    public long CacheMisses => _cacheMisses;
    public long RecalcTimeMs => _recalcTimeMs;
    public double CacheHitRate => (_cacheHits + _cacheMisses) == 0 ? 0
        : (double)_cacheHits / (_cacheHits + _cacheMisses);

    public void RecordParse() => Interlocked.Increment(ref _parseCount);
    public void RecordEval() => Interlocked.Increment(ref _evalCount);
    public void RecordCacheHit() => Interlocked.Increment(ref _cacheHits);
    public void RecordCacheMiss() => Interlocked.Increment(ref _cacheMisses);
    public void RecordRecalcTime(long ms) => Interlocked.Add(ref _recalcTimeMs, ms);

    public void Reset()
    {
        _parseCount = 0; _evalCount = 0;
        _cacheHits = 0; _cacheMisses = 0;
        _recalcTimeMs = 0;
    }

    public override string ToString() =>
        $"Parse={ParseCount}, Eval={EvalCount}, CacheHit={CacheHitRate:P1}, RecalcMs={RecalcTimeMs}";
}
