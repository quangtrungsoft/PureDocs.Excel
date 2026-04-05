namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Thrown when a formula needs a precedent that hasn't been evaluated yet.
/// Used by the dynamic reordering mechanism for INDIRECT, OFFSET, and other
/// functions that create runtime dependencies unknown at parse time.
/// </summary>
public sealed class PrecedentNotReadyException : Exception
{
    public CellAddress NeededCell { get; }

    public PrecedentNotReadyException(CellAddress cell)
        : base($"Precedent {cell} not yet evaluated")
    {
        NeededCell = cell;
    }
}

/// <summary>
/// Evaluator wrapper with dynamic reordering support.
/// When a formula references a cell via INDIRECT/OFFSET that hasn't been evaluated,
/// it defers evaluation and re-queues after the precedent is ready.
///
/// Algorithm:
///   1. Start with static topo order
///   2. For each cell, try to evaluate
///   3. If PrecedentNotReadyException → evaluate needed precedent recursively
///   4. Retry cell after precedent is ready
///   5. Track dependency chain to detect dynamic circular references
///   6. Max depth limit prevents stack overflow
/// </summary>
public sealed class DynamicEvaluator
{
    private readonly CalcChain _chain;
    private const int MaxRecursionDepth = 100;

    public DynamicEvaluator(CalcChain chain)
    {
        _chain = chain;
    }

    /// <summary>
    /// Evaluates a list of cells with dynamic reordering support.
    /// Uses recursive evaluation with cycle detection to handle dynamic dependencies.
    /// </summary>
    public int Evaluate(List<CellAddress> order, Worksheet worksheet)
    {
        int count = 0;
        var evaluated = new HashSet<CellAddress>();
        var failed = new HashSet<CellAddress>(); // Cells that failed permanently
        
        // Single pass with recursive dependency resolution
        foreach (var cell in order)
        {
            if (evaluated.Contains(cell) || failed.Contains(cell))
                continue;
                
            var evaluating = new HashSet<CellAddress>(); // Track current evaluation chain for cycle detection
            if (TryEvaluateCellRecursive(cell, worksheet, evaluated, failed, evaluating, 0))
                count++;
        }

        return count;
    }

    /// <summary>
    /// Recursively evaluates a cell, resolving dynamic precedents as needed.
    /// </summary>
    /// <param name="cell">Cell to evaluate</param>
    /// <param name="worksheet">Target worksheet</param>
    /// <param name="evaluated">Set of successfully evaluated cells</param>
    /// <param name="failed">Set of cells that failed permanently (circular/error)</param>
    /// <param name="evaluating">Current evaluation chain for cycle detection</param>
    /// <param name="depth">Current recursion depth</param>
    /// <returns>True if evaluation succeeded</returns>
    private bool TryEvaluateCellRecursive(CellAddress cell, Worksheet worksheet,
        HashSet<CellAddress> evaluated, HashSet<CellAddress> failed,
        HashSet<CellAddress> evaluating, int depth)
    {
        // Already evaluated or failed
        if (evaluated.Contains(cell)) return true;
        if (failed.Contains(cell)) return false;
        
        // Check for dynamic circular reference
        if (evaluating.Contains(cell))
        {
            // Circular dependency via INDIRECT/OFFSET detected
            _chain.SetCachedValueDirect(cell, FormulaValue.ErrorRef);
            failed.Add(cell);
            return false;
        }
        
        // Check recursion depth to prevent stack overflow
        if (depth >= MaxRecursionDepth)
        {
            _chain.SetCachedValueDirect(cell, FormulaValue.ErrorRef);
            failed.Add(cell);
            return false;
        }

        evaluating.Add(cell);
        
        try
        {
            var ast = _chain.GetCachedAst(cell);
            if (ast == null)
            {
                evaluated.Add(cell);
                return true; // No formula = success
            }

            var context = new FormulaContext(worksheet, cell.Row, cell.Column);
            var result = ast.Evaluate(context);
            _chain.SetCachedValueDirect(cell, result);
            evaluated.Add(cell);
            return true;
        }
        catch (PrecedentNotReadyException ex)
        {
            // Try to evaluate the needed precedent first
            if (TryEvaluateCellRecursive(ex.NeededCell, worksheet, evaluated, failed, evaluating, depth + 1))
            {
                // Precedent ready, retry this cell
                try
                {
                    var ast = _chain.GetCachedAst(cell);
                    if (ast == null)
                    {
                        evaluated.Add(cell);
                        return true;
                    }

                    var context = new FormulaContext(worksheet, cell.Row, cell.Column);
                    var result = ast.Evaluate(context);
                    _chain.SetCachedValueDirect(cell, result);
                    evaluated.Add(cell);
                    return true;
                }
                catch
                {
                    // Still failing after precedent was evaluated - mark as error
                    _chain.SetCachedValueDirect(cell, FormulaValue.ErrorRef);
                    failed.Add(cell);
                    return false;
                }
            }
            else
            {
                // Precedent failed - this cell fails too
                _chain.SetCachedValueDirect(cell, FormulaValue.ErrorRef);
                failed.Add(cell);
                return false;
            }
        }
        catch (Exception)
        {
            _chain.SetCachedValueDirect(cell, FormulaValue.ErrorValue);
            failed.Add(cell);
            return false;
        }
        finally
        {
            evaluating.Remove(cell);
        }
    }
}
