using System.Collections.Concurrent;
using System.Diagnostics;

namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Manages formula calculation chain for a worksheet.
/// V2: RangeSet-based dependencies, parallel level recalc, correct cycle reporting.
/// V3: Optional LRU cache mode for memory-capped large workbooks.
/// V4: Thread-safe dirty cell tracking for parallel operations.
/// </summary>
public sealed class CalcChain
{
    private readonly DependencyGraph _graph = new();
    private readonly ConcurrentDictionary<CellAddress, byte> _dirtyCells = new(); // Thread-safe
    private readonly object _dirtyLock = new(); // Lock for MarkDirty traversal
    private readonly ConcurrentDictionary<CellAddress, FormulaNode> _cellAstCache = new();
    private readonly ConcurrentDictionary<CellAddress, FormulaValue> _valueCache = new();
    private LruCache<CellAddress, FormulaValue>? _lruValueCache;
    private bool _fullRecalcNeeded = true;

    /// <summary>Underlying dependency graph.</summary>
    public DependencyGraph Graph => _graph;

    /// <summary>Number of dirty cells awaiting recalc.</summary>
    public int DirtyCount => _dirtyCells.Count;

    /// <summary>Performance counters.</summary>
    public FormulaPerformanceCounters Counters { get; } = new();

    /// <summary>
    /// Enables LRU cache mode with a maximum capacity. When enabled, the value cache
    /// will evict least-recently-used values to stay under the memory cap.
    /// Recommended for workbooks with >50,000 formula cells.
    /// </summary>
    /// <param name="capacity">Maximum number of cached values (0 to disable LRU mode)</param>
    public void EnableLruCache(int capacity)
    {
        if (capacity <= 0)
        {
            _lruValueCache = null;
        }
        else
        {
            _lruValueCache = new LruCache<CellAddress, FormulaValue>(capacity);
            // Migrate existing cache entries
            foreach (var kv in _valueCache)
                _lruValueCache.Set(kv.Key, kv.Value);
            _valueCache.Clear();
        }
    }

    /// <summary>Gets LRU cache statistics if LRU mode is enabled.</summary>
    public (long hits, long misses, long evictions, double hitRate)? GetLruStats()
    {
        if (_lruValueCache == null) return null;
        return (_lruValueCache.Hits, _lruValueCache.Misses, 
                _lruValueCache.Evictions, _lruValueCache.HitRate);
    }

    /// <summary>
    /// Registers or updates a formula for a cell. Parses it, extracts dependencies,
    /// and marks affected cells dirty.
    /// </summary>
    public void SetFormula(CellAddress cell, string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            RemoveFormula(cell);
            return;
        }

        formula = formula.TrimStart('=');

        // Parse and cache AST
        var lexer = new FormulaLexer(formula);
        var tokens = lexer.Tokenize();
        var ast = new FormulaParser(tokens).Parse();
        _cellAstCache[cell] = ast;

        // Extract dependencies using compact RangeSet (area-based)
        var rangeSet = DependencyCollector.CollectRangeSet(ast);
        bool isVolatile = DependencyCollector.ContainsVolatile(ast);

        _graph.SetDependencies(cell, rangeSet, isVolatile);

        // Mark this cell and all its dependents dirty
        MarkDirty(cell);
    }

    /// <summary>Removes formula tracking for a cell.</summary>
    public void RemoveFormula(CellAddress cell)
    {
        _cellAstCache.TryRemove(cell, out _);
        RemoveCachedValue(cell);
        Spills.ClearSpill(cell);
        _graph.RemoveDependencies(cell);
        MarkDirty(cell);
    }

    /// <summary>
    /// Marks a cell and all its dependents as needing recalc.
    /// Thread-safe: can be called from multiple threads during parallel operations.
    /// </summary>
    public void MarkDirty(CellAddress cell)
    {
        // Lock to prevent concurrent traversal issues
        // Multiple threads can call this simultaneously when SetValue is called during parallel recalc
        lock (_dirtyLock)
        {
            _dirtyCells.TryAdd(cell, 0);
            RemoveCachedValue(cell);

            var visited = new HashSet<CellAddress>();
            var queue = new Queue<CellAddress>();
            queue.Enqueue(cell);

            while (queue.Count > 0)
            {
                var current = queue.Dequeue();
                foreach (var dep in _graph.GetAllDependents(current))
                {
                    if (visited.Add(dep))
                    {
                        _dirtyCells.TryAdd(dep, 0);
                        RemoveCachedValue(dep);
                        queue.Enqueue(dep);
                    }
                }
            }
        }
    }

    /// <summary>Marks all formula cells as dirty (full recalc).</summary>
    public void MarkAllDirty()
    {
        _fullRecalcNeeded = true;
        _valueCache.Clear();
    }

    /// <summary>
    /// Recalculates only dirty cells in correct topological order.
    /// Reports cycles precisely. Returns the number of cells recalculated.
    /// </summary>
    public int Recalculate(Worksheet worksheet)
    {
        var sw = Stopwatch.StartNew();

        try
        {
            if (_fullRecalcNeeded)
            {
                _fullRecalcNeeded = false;
                return FullRecalculate(worksheet);
            }

            if (_dirtyCells.IsEmpty)
                return 0;

            // Convert to HashSet for GetRecalcOrderWithCycles
            var dirtyCellsSnapshot = new HashSet<CellAddress>(_dirtyCells.Keys);
            var (order, cycleCells) = _graph.GetRecalcOrderWithCycles(dirtyCellsSnapshot);

            // Handle cycles: mark exactly which cells are cyclic
            // NOTE: Excel supports iterative calculation for circular refs, but we treat them as errors.
            // To enable iterative calc, set EnableIterativeCalculation = true (not yet implemented).
            if (cycleCells != null)
            {
                foreach (var cell in cycleCells)
                    SetCachedValue(cell, FormulaValue.Error(FormulaError.Ref));
            }

            // Evaluate non-cyclic cells in topo order
            int recalcCount = EvaluateInOrderConverging(order, worksheet);

            _dirtyCells.Clear();
            return recalcCount + (cycleCells?.Count ?? 0);
        }
        finally
        {
            sw.Stop();
            Counters.RecordRecalcTime(sw.ElapsedMilliseconds);
        }
    }

    /// <summary>
    /// Parallel recalculate: evaluates independent cells in parallel using level-based
    /// parallelism. Cells at the same topological level have no dependencies on each other.
    /// </summary>
    /// <remarks>
    /// Uses a shared thread-safe evaluatingCells set across all parallel evaluations
    /// to properly detect circular references even in parallel mode.
    /// </remarks>
    public int ParallelRecalculate(Worksheet worksheet)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            if (_fullRecalcNeeded)
            {
                _fullRecalcNeeded = false;
                BuildFromWorksheet(worksheet);
            }

            if (_dirtyCells.IsEmpty)
                return 0;

            // Convert to HashSet for GetRecalcOrderWithCycles
            var dirtyCellsSnapshot = new HashSet<CellAddress>(_dirtyCells.Keys);
            var (order, cycleCells) = _graph.GetRecalcOrderWithCycles(dirtyCellsSnapshot);

            if (cycleCells != null)
                foreach (var cell in cycleCells)
                    SetCachedValue(cell, FormulaValue.Error(FormulaError.Ref));

            // Group into levels for parallel execution
            var levels = BuildLevels(order);
            int count = 0;

            // Shared thread-safe set for circular reference detection across all parallel threads
            var sharedEvaluatingCells = new ConcurrentDictionary<string, byte>();

            foreach (var level in levels)
            {
                if (level.Count <= 8)
                {
                    // Sequential for small levels (avoid thread overhead)
                    foreach (var cell in level)
                        EvaluateSingleCellWithSharedContext(cell, worksheet, sharedEvaluatingCells);
                }
                else
                {
                    // Parallel for larger levels
                    Parallel.ForEach(level, cell => 
                        EvaluateSingleCellWithSharedContext(cell, worksheet, sharedEvaluatingCells));
                }
                count += level.Count;
            }

            _dirtyCells.Clear();
            return count + (cycleCells?.Count ?? 0);
        }
        finally
        {
            sw.Stop();
            Counters.RecordRecalcTime(sw.ElapsedMilliseconds);
        }
    }

    /// <summary>Groups topologically sorted cells into parallel levels.</summary>
    private List<List<CellAddress>> BuildLevels(List<CellAddress> topoOrder)
    {
        var levels = new List<List<CellAddress>>();
        var cellLevel = new Dictionary<CellAddress, int>();
        var topoSet = new HashSet<CellAddress>(topoOrder);

        foreach (var cell in topoOrder)
        {
            int maxPrecLevel = -1;
            foreach (var prec in _graph.GetPrecedents(cell))
            {
                if (topoSet.Contains(prec) && cellLevel.TryGetValue(prec, out int precLevel))
                    maxPrecLevel = Math.Max(maxPrecLevel, precLevel);
            }

            int myLevel = maxPrecLevel + 1;
            cellLevel[cell] = myLevel;

            while (levels.Count <= myLevel)
                levels.Add(new List<CellAddress>());
            levels[myLevel].Add(cell);
        }

        return levels;
    }

    /// <summary>
    /// Evaluates <paramref name="order"/>, repeating the pass while spill regions keep changing.
    /// A cell that reads a spilled value depends on the spill anchor, but that dependency is not
    /// tracked in the static graph, so a first pass may evaluate the reader before the anchor
    /// spills. Re-running once the spill layout is stable converges those readers (bounded to a
    /// few passes to guard against oscillation).
    /// </summary>
    private int EvaluateInOrderConverging(List<CellAddress> order, Worksheet worksheet)
    {
        const int maxPasses = 4;
        int count = 0;
        for (int pass = 0; pass < maxPasses; pass++)
        {
            long before = Spills.ChangeCount;
            count = EvaluateInOrder(order, worksheet);
            if (Spills.ChangeCount == before) break; // spill layout stable → done
        }
        return count;
    }

    private int EvaluateInOrder(List<CellAddress> order, Worksheet worksheet)
    {
        int count = 0;
        var context = new FormulaContext(worksheet, namedRanges: worksheet.NamedRanges);
        foreach (var cell in order)
        {
            if (_cellAstCache.TryGetValue(cell, out var ast))
            {
                // Set formula position for implicit intersection (@) operator
                context.FormulaRow = cell.Row;
                context.FormulaCol = cell.Column;
                
                var result = ast.Evaluate(context);
                result = ApplySpill(cell, result, worksheet);
                SetCachedValue(cell, result);
                count++;
                Counters.RecordEval();
            }
        }
        return count;
    }

    private void EvaluateSingleCell(CellAddress cell, Worksheet worksheet)
    {
        if (_cellAstCache.TryGetValue(cell, out var ast))
        {
            // Create context with formula position for implicit intersection
            var context = new FormulaContext(worksheet, cell.Row, cell.Column,
                namedRanges: worksheet.NamedRanges);
            var result = ast.Evaluate(context);
            result = ApplySpill(cell, result, worksheet);
            SetCachedValue(cell, result);
            Counters.RecordEval();
        }
    }

    /// <summary>
    /// Evaluates a single cell using a shared thread-safe evaluatingCells set.
    /// Used by ParallelRecalculate for proper cross-thread circular reference detection.
    /// </summary>
    private void EvaluateSingleCellWithSharedContext(CellAddress cell, Worksheet worksheet, 
        ConcurrentDictionary<string, byte> sharedEvaluatingCells)
    {
        if (_cellAstCache.TryGetValue(cell, out var ast))
        {
            var context = new FormulaContext(worksheet, sharedEvaluatingCells, worksheet.NamedRanges);
            // Set formula position for implicit intersection (@) operator
            context.FormulaRow = cell.Row;
            context.FormulaCol = cell.Column;
            
            var result = ast.Evaluate(context);
            result = ApplySpill(cell, result, worksheet);
            SetCachedValue(cell, result);
            Counters.RecordEval();
        }
    }

    /// <summary>
    /// Handles dynamic-array spill for a formula result. A multi-cell array result
    /// spills into the cells below/right of the anchor; the anchor caches the top-left
    /// value. Returns <c>#SPILL!</c> if the target region is not empty. Scalars and
    /// 1×1 arrays are returned unchanged (unwrapped).
    /// </summary>
    private FormulaValue ApplySpill(CellAddress anchor, FormulaValue result, Worksheet worksheet)
    {
        // Refresh: drop any prior spill from this anchor before re-evaluating.
        Spills.ClearSpill(anchor);

        if (!result.IsArray) return result;
        var arr = result.ArrayVal;
        if (arr.Rows <= 1 && arr.Columns <= 1)
            return arr.Length > 0 ? arr[0] : FormulaValue.Blank;

        var region = Spills.TrySpill(anchor, arr,
            target => IsCellSpillEmpty(target, anchor, worksheet));
        return region == null
            ? FormulaValue.Error(FormulaError.Spill) // blocked by existing content
            : arr[0, 0];
    }

    /// <summary>
    /// A target cell is available for spill if it is the anchor itself, has no own content,
    /// or is a cell that a previously-saved dynamic array spilled into (tracked in
    /// <see cref="_spillOwnedOnLoad"/>) — so re-opening and recalculating does not falsely
    /// report <c>#SPILL!</c> against the spill's own cached values.
    /// </summary>
    private bool IsCellSpillEmpty(CellAddress target, CellAddress anchor, Worksheet worksheet)
    {
        if (target.Equals(anchor)) return true;
        if (_spillOwnedOnLoad.Contains(target)) return true;
        var cell = worksheet.TryGetCell(target.ToReference());
        return cell == null || (!cell.HasFormula && cell.GetValue() == null);
    }

    /// <summary>
    /// Cells that were spilled into by a dynamic array in the file being read (i.e. they lie
    /// within an array formula's <c>ref</c> but hold no formula of their own). Recorded during
    /// <see cref="BuildFromWorksheet"/> so they are treated as spill-available, not user content.
    /// </summary>
    private readonly HashSet<CellAddress> _spillOwnedOnLoad = new();

    /// <summary>Full recalculate of all formula cells.</summary>
    private int FullRecalculate(Worksheet worksheet)
    {
        BuildFromWorksheet(worksheet);

        var allFormulaCells = new List<CellAddress>();
        foreach (var kv in _cellAstCache)
            allFormulaCells.Add(kv.Key);

        var (order, cycleCells) = _graph.GetRecalcOrderWithCycles(allFormulaCells);

        if (cycleCells != null)
            foreach (var cell in cycleCells)
                SetCachedValue(cell, FormulaValue.Error(FormulaError.Ref));

        int count = EvaluateInOrderConverging(order, worksheet);
        _dirtyCells.Clear();
        return count + (cycleCells?.Count ?? 0);
    }

    /// <summary>Builds dependency graph from worksheet's current formulas.</summary>
    public void BuildFromWorksheet(Worksheet worksheet)
    {
        _graph.Clear();
        _cellAstCache.Clear();
        _valueCache.Clear();
        _spillOwnedOnLoad.Clear();

        var sheetData = worksheet.GetSheetData();
        if (sheetData == null) return;

        foreach (var row in sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>())
        {
            foreach (var oxCell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                if (oxCell.CellFormula != null && !string.IsNullOrEmpty(oxCell.CellFormula.Text))
                {
                    string cellRef = oxCell.CellReference?.Value ?? "";
                    if (string.IsNullOrEmpty(cellRef)) continue;

                    var addr = CellAddress.FromReference(cellRef);
                    string formula = oxCell.CellFormula.Text.TrimStart('=');

                    // An array formula with a multi-cell ref is a dynamic-array anchor;
                    // record the cells it spilled into so they aren't seen as blocking content.
                    RecordSpillOwnership(oxCell, addr);

                    try
                    {
                        var lexer = new FormulaLexer(formula);
                        var tokens = lexer.Tokenize();
                        var ast = new FormulaParser(tokens).Parse();
                        _cellAstCache[addr] = ast;
                        Counters.RecordParse();

                        // Use compact RangeSet (area-based dependencies)
                        var rangeSet = DependencyCollector.CollectRangeSet(ast);
                        bool isVolatile = DependencyCollector.ContainsVolatile(ast);
                        _graph.SetDependencies(addr, rangeSet, isVolatile);
                    }
                    catch
                    {
                        // Skip unparseable formulas
                    }
                }
            }
        }
    }

    /// <summary>
    /// If the given OpenXML cell is a dynamic-array anchor (array formula with a multi-cell
    /// <c>ref</c>), records every non-anchor cell inside that ref as spill-owned.
    /// </summary>
    private void RecordSpillOwnership(DocumentFormat.OpenXml.Spreadsheet.Cell oxCell, CellAddress anchor)
    {
        var f = oxCell.CellFormula;
        if (f?.FormulaType == null || f.FormulaType.Value != DocumentFormat.OpenXml.Spreadsheet.CellFormulaValues.Array)
            return;
        string? refRange = f.Reference?.Value;
        if (string.IsNullOrEmpty(refRange)) return;

        int colon = refRange.IndexOf(':');
        if (colon <= 0) return; // single-cell ref → not a spill
        try
        {
            CellReference.Parse(refRange[..colon], out int sr, out int sc);
            CellReference.Parse(refRange[(colon + 1)..], out int er, out int ec);
            for (int r = sr; r <= er; r++)
                for (int c = sc; c <= ec; c++)
                {
                    var cell = new CellAddress(r, c, anchor.SheetIndex);
                    if (!cell.Equals(anchor)) _spillOwnedOnLoad.Add(cell);
                }
        }
        catch { /* malformed ref → ignore */ }
    }

    /// <summary>Gets the cached value for a cell, or null if not calculated.</summary>
    public FormulaValue? GetCachedValue(CellAddress cell)
    {
        if (_lruValueCache != null)
            return _lruValueCache.TryGetValue(cell, out var v) ? v : null;
        return _valueCache.TryGetValue(cell, out var val) ? val : null;
    }

    /// <summary>Clears all cached data.</summary>
    public void Clear()
    {
        _graph.Clear();
        _dirtyCells.Clear();
        _cellAstCache.Clear();
        _valueCache.Clear();
        _lruValueCache?.Clear();
        _fullRecalcNeeded = true;
    }

    /// <summary>Diagnostics: number of cached ASTs.</summary>
    public int CachedAstCount => _cellAstCache.Count;

    /// <summary>Diagnostics: number of cached values.</summary>
    public int CachedValueCount => _lruValueCache?.Count ?? _valueCache.Count;

    /// <summary>Spill engine for dynamic array support.</summary>
    public SpillEngine Spills { get; } = new();

    // ── Cache Helper Methods ──────────────────────────────────────────

    /// <summary>Sets a cached value, using LRU cache if enabled.</summary>
    private void SetCachedValue(CellAddress cell, FormulaValue value)
    {
        if (_lruValueCache != null)
            _lruValueCache.Set(cell, value);
        else
            _valueCache[cell] = value;
    }

    /// <summary>Removes a cached value, using LRU cache if enabled.</summary>
    private void RemoveCachedValue(CellAddress cell)
    {
        if (_lruValueCache != null)
            _lruValueCache.Remove(cell);
        else
            _valueCache.TryRemove(cell, out _);
    }
}
