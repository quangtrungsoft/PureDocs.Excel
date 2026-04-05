namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Directed acyclic graph tracking formula dependencies.
/// V2: Uses RangeSet for compact area-based storage + improved cycle reporting.
/// 
/// Terminology:
///   - Precedent: a cell that this formula depends on (A1 is precedent of =A1+1)
///   - Dependent: a cell whose formula references this cell (=A1+1 is dependent of A1)
/// 
/// When cell X changes → find all dependents of X → mark dirty → recalc in topo order.
/// </summary>
public sealed class DependencyGraph
{
    // Cell → its precedents as compact RangeSet (cells + areas it reads from)
    private readonly Dictionary<CellAddress, RangeSet> _precedentsCompact = new();
    // Cell → its precedents expanded (for backward compat with topo sort)
    private readonly Dictionary<CellAddress, HashSet<CellAddress>> _precedents = new();
    // Cell → its dependents (cells that read from it)
    private readonly Dictionary<CellAddress, HashSet<CellAddress>> _dependents = new();
    // Cells containing volatile formulas (always recalc)
    private readonly HashSet<CellAddress> _volatileCells = new();
    // All formula cells
    private readonly HashSet<CellAddress> _formulaCells = new();

    /// <summary>Number of tracked formula cells.</summary>
    public int FormulaCellCount => _formulaCells.Count;

    /// <summary>Number of volatile formula cells.</summary>
    public int VolatileCellCount => _volatileCells.Count;

    /// <summary>
    /// Sets the dependencies for a formula cell using compact RangeSet.
    /// This is the preferred method — stores areas without expanding.
    /// </summary>
    public void SetDependencies(CellAddress cell, RangeSet rangeSet, bool isVolatile)
    {
        // Remove old dependencies first
        RemoveDependencies(cell);

        _formulaCells.Add(cell);
        _precedentsCompact[cell] = rangeSet;

        // For topo sort: we still need expanded precedents of individual cell refs
        // but NOT range areas (those are checked via RangeSet.Contains)
        var expandedPrecs = new HashSet<CellAddress>(rangeSet.Singles);
        _precedents[cell] = expandedPrecs;

        // Register reverse (dependent) links for single cells
        foreach (var prec in rangeSet.Singles)
        {
            if (!_dependents.TryGetValue(prec, out var deps))
            {
                deps = new HashSet<CellAddress>();
                _dependents[prec] = deps;
            }
            deps.Add(cell);
        }

        if (isVolatile)
            _volatileCells.Add(cell);
        else
            _volatileCells.Remove(cell);
    }

    /// <summary>
    /// Legacy: Sets dependencies using expanded HashSet (backward compat).
    /// </summary>
    public void SetDependencies(CellAddress cell, HashSet<CellAddress> newPrecedents, bool isVolatile)
    {
        RemoveDependencies(cell);

        _formulaCells.Add(cell);
        _precedents[cell] = newPrecedents;

        foreach (var prec in newPrecedents)
        {
            if (!_dependents.TryGetValue(prec, out var deps))
            {
                deps = new HashSet<CellAddress>();
                _dependents[prec] = deps;
            }
            deps.Add(cell);
        }

        if (isVolatile)
            _volatileCells.Add(cell);
        else
            _volatileCells.Remove(cell);
    }

    /// <summary>Removes all dependency tracking for a cell.</summary>
    public void RemoveDependencies(CellAddress cell)
    {
        if (_precedents.TryGetValue(cell, out var oldPrecs))
        {
            foreach (var prec in oldPrecs)
                if (_dependents.TryGetValue(prec, out var deps))
                {
                    deps.Remove(cell);
                    if (deps.Count == 0) _dependents.Remove(prec);
                }
            _precedents.Remove(cell);
        }
        _precedentsCompact.Remove(cell);
        _formulaCells.Remove(cell);
        _volatileCells.Remove(cell);
    }

    /// <summary>Gets direct dependents of a cell.</summary>
    public IReadOnlySet<CellAddress> GetDependents(CellAddress cell)
    {
        return _dependents.TryGetValue(cell, out var deps) ? deps : EmptySet;
    }

    /// <summary>
    /// Gets ALL dependents of a cell, including formulas that reference ranges containing this cell.
    /// This is the correct method for dirty marking.
    /// </summary>
    public HashSet<CellAddress> GetAllDependents(CellAddress cell)
    {
        var result = new HashSet<CellAddress>();

        // Direct single-cell dependents (fast, O(1))
        if (_dependents.TryGetValue(cell, out var directDeps))
            foreach (var dep in directDeps)
                result.Add(dep);

        // Area dependents: scan formula cells whose RangeSet areas contain this cell
        // This is O(formula_cells × areas_per_formula), typically small
        // NOTE: We intentionally don't skip cells already in result because:
        // 1. HashSet.Add handles duplicates automatically
        // 2. A cell could be both a direct and area dependent (different formulas)
        foreach (var kv in _precedentsCompact)
        {
            foreach (var area in kv.Value.Areas)
            {
                if (area.Contains(cell))
                {
                    result.Add(kv.Key);
                    break; // Found in one area, no need to check other areas for same cell
                }
            }
        }

        return result;
    }

    /// <summary>Gets direct precedents of a cell.</summary>
    public IReadOnlySet<CellAddress> GetPrecedents(CellAddress cell)
    {
        return _precedents.TryGetValue(cell, out var precs) ? precs : EmptySet;
    }

    /// <summary>Gets all volatile cells.</summary>
    public IReadOnlySet<CellAddress> VolatileCells => _volatileCells;

    /// <summary>
    /// Given a set of changed cells, returns ALL cells that need recalculation
    /// in correct topological order (Kahn's algorithm).
    /// Also includes volatile cells.
    /// Returns (orderedCells, cycleCells) — cycleCells is null if no cycles.
    /// </summary>
    public (List<CellAddress> order, HashSet<CellAddress>? cycleCells) GetRecalcOrderWithCycles(
        IEnumerable<CellAddress> changedCells)
    {
        // Step 1: BFS to find all affected formula cells
        var affected = new HashSet<CellAddress>();
        var queue = new Queue<CellAddress>();

        // Seed with dependents of changed cells (using area-aware lookup)
        foreach (var cell in changedCells)
        {
            var deps = GetAllDependents(cell);
            foreach (var dep in deps)
                if (affected.Add(dep)) queue.Enqueue(dep);
        }
        // Volatile cells always recalc
        foreach (var vol in _volatileCells)
            if (affected.Add(vol)) queue.Enqueue(vol);

        // BFS: transitively find all dependents
        while (queue.Count > 0)
        {
            var cell = queue.Dequeue();
            var deps = GetAllDependents(cell);
            foreach (var dep in deps)
                if (affected.Add(dep)) queue.Enqueue(dep);
        }

        if (affected.Count == 0)
            return (new List<CellAddress>(), null);

        // Step 2: Topological sort (Kahn's algorithm) on affected cells
        // IMPORTANT: Must consider BOTH single cell precedents AND area precedents
        var inDegree = new Dictionary<CellAddress, int>();
        foreach (var cell in affected)
            inDegree[cell] = 0;

        foreach (var cell in affected)
        {
            // Count single cell precedents
            if (_precedents.TryGetValue(cell, out var precs))
                foreach (var prec in precs)
                    if (affected.Contains(prec))
                        inDegree[cell]++;

            // Count area precedents: if this cell's formula depends on a range,
            // and any cell in that range is in 'affected', increment inDegree
            if (_precedentsCompact.TryGetValue(cell, out var rangeSet))
            {
                foreach (var area in rangeSet.Areas)
                {
                    // Check if any affected cell falls within this area
                    foreach (var affectedCell in affected)
                    {
                        // Don't count self-dependency here
                        if (affectedCell.Equals(cell)) continue;
                        
                        if (area.Contains(affectedCell))
                        {
                            inDegree[cell]++;
                            // Note: We count each affected cell in the area as a dependency
                            // This ensures proper ordering when multiple cells in a range change
                        }
                    }
                }
            }
        }

        var topoQueue = new Queue<CellAddress>();
        foreach (var kv in inDegree)
            if (kv.Value == 0) topoQueue.Enqueue(kv.Key);

        var result = new List<CellAddress>(affected.Count);
        while (topoQueue.Count > 0)
        {
            var cell = topoQueue.Dequeue();
            result.Add(cell);

            var deps = GetAllDependents(cell);
            foreach (var dep in deps)
                if (inDegree.ContainsKey(dep))
                {
                    inDegree[dep]--;
                    if (inDegree[dep] == 0) topoQueue.Enqueue(dep);
                }
        }

        // Cycle detection: extract exactly which cells are in cycles
        if (result.Count < affected.Count)
        {
            var cycleCells = new HashSet<CellAddress>(
                affected.Where(c => !result.Contains(c)));
            return (result, cycleCells);
        }

        return (result, null);
    }

    /// <summary>Legacy: returns null on cycle (backward compat).</summary>
    public List<CellAddress>? GetRecalcOrder(IEnumerable<CellAddress> changedCells)
    {
        var (order, cycles) = GetRecalcOrderWithCycles(changedCells);
        return cycles != null ? null : order;
    }

    /// <summary>Clears all tracking data.</summary>
    public void Clear()
    {
        _precedents.Clear();
        _precedentsCompact.Clear();
        _dependents.Clear();
        _volatileCells.Clear();
        _formulaCells.Clear();
    }

    private static readonly HashSet<CellAddress> EmptySet = new();
}
