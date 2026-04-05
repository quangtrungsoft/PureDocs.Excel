namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Manages dynamic array (spill) behavior for formulas that return arrays.
/// Tracks spill ranges, detects spill blocking, and manages the spill anchor → region mapping.
///
/// Excel 365+ semantics:
///   =SORT(A1:A10) in C1 → spills results into C1:C10
///   If D1:D10 already has data → #SPILL! in C1
/// </summary>
/// <remarks>
/// <para>
/// <strong>STATUS: WORK IN PROGRESS</strong>
/// </para>
/// <para>
/// This engine is implemented but NOT YET INTEGRATED into the main formula evaluation pipeline.
/// The following integration points are needed:
/// </para>
/// <list type="bullet">
///   <item>CalcChain.EvaluateInOrder: After evaluating a formula that returns an array,
///         call SpillEngine.TrySpill() and handle #SPILL! if blocked.</item>
///   <item>FormulaContext.GetCellValue: Check SpillEngine.GetSpillValue() for cells that
///         are part of a spill region but don't contain their own formula.</item>
///   <item>Worksheet.SetValue: Call SpillEngine.ClearSpill() if user writes to a spill anchor,
///         and check for spill blocking if writing to a cell in a spill region.</item>
///   <item>CalcChain.MarkDirty: When a spill anchor is dirtied, also dirty the spill region.</item>
/// </list>
/// <para>
/// To enable spill functionality, integrate the SpillEngine from CalcChain.Spills into
/// the evaluation pipeline using the above integration points.
/// </para>
/// </remarks>
public sealed class SpillEngine
{
    /// <summary>Tracks which cells are occupied by spill output.</summary>
    private readonly Dictionary<CellAddress, SpillRegion> _spillAnchors = new();
    /// <summary>Reverse map: occupied cell → anchor cell.</summary>
    private readonly Dictionary<CellAddress, CellAddress> _occupiedBySpill = new();

    /// <summary>Information about a spill region.</summary>
    public sealed class SpillRegion
    {
        public CellAddress Anchor { get; init; }
        public int Rows { get; init; }
        public int Columns { get; init; }
        public ArrayValue Data { get; init; } = ArrayValue.Empty;
    }

    /// <summary>
    /// Attempts to spill an array result from a formula cell.
    /// Returns the SpillRegion on success, or null if blocked.
    /// </summary>
    public SpillRegion? TrySpill(CellAddress anchor, ArrayValue array,
        Func<CellAddress, bool> isCellEmpty)
    {
        // Remove previous spill from this anchor
        ClearSpill(anchor);

        if (array.Rows <= 1 && array.Columns <= 1)
        {
            // Scalar or 1×1 — no spill needed
            return null;
        }

        // Check if all target cells are available
        for (int r = 0; r < array.Rows; r++)
            for (int c = 0; c < array.Columns; c++)
            {
                if (r == 0 && c == 0) continue; // anchor cell itself
                var target = new CellAddress(anchor.Row + r, anchor.Column + c, anchor.SheetIndex);

                // Blocked by another formula/data cell or another spill
                if (!isCellEmpty(target) ||
                    (_occupiedBySpill.TryGetValue(target, out var otherAnchor) && otherAnchor != anchor))
                {
                    return null; // Spill blocked → caller should return #SPILL!
                }
            }

        // Register spill region
        var region = new SpillRegion
        {
            Anchor = anchor,
            Rows = array.Rows,
            Columns = array.Columns,
            Data = array
        };
        _spillAnchors[anchor] = region;

        // Mark all cells as occupied
        for (int r = 0; r < array.Rows; r++)
            for (int c = 0; c < array.Columns; c++)
            {
                var cell = new CellAddress(anchor.Row + r, anchor.Column + c, anchor.SheetIndex);
                _occupiedBySpill[cell] = anchor;
            }

        return region;
    }

    /// <summary>Clears a spill region for an anchor cell.</summary>
    public void ClearSpill(CellAddress anchor)
    {
        if (!_spillAnchors.TryGetValue(anchor, out var region)) return;

        for (int r = 0; r < region.Rows; r++)
            for (int c = 0; c < region.Columns; c++)
            {
                var cell = new CellAddress(anchor.Row + r, anchor.Column + c, anchor.SheetIndex);
                _occupiedBySpill.Remove(cell);
            }
        _spillAnchors.Remove(anchor);
    }

    /// <summary>Gets the spill value at a cell (if it's part of a spill region).</summary>
    public FormulaValue? GetSpillValue(CellAddress cell)
    {
        if (!_occupiedBySpill.TryGetValue(cell, out var anchor)) return null;
        if (!_spillAnchors.TryGetValue(anchor, out var region)) return null;

        int r = cell.Row - anchor.Row;
        int c = cell.Column - anchor.Column;
        return region.Data[r, c];
    }

    /// <summary>Checks if a cell is part of any spill region.</summary>
    public bool IsSpillOccupied(CellAddress cell) => _occupiedBySpill.ContainsKey(cell);

    /// <summary>Gets the spill anchor for a cell (if it's in a spill region).</summary>
    public CellAddress? GetSpillAnchor(CellAddress cell)
        => _occupiedBySpill.TryGetValue(cell, out var anchor) ? anchor : null;

    /// <summary>Gets the spill region for an anchor cell.</summary>
    public SpillRegion? GetRegion(CellAddress anchor)
        => _spillAnchors.TryGetValue(anchor, out var r) ? r : null;

    /// <summary>Number of active spill regions.</summary>
    public int RegionCount => _spillAnchors.Count;

    /// <summary>Clears all spill data.</summary>
    public void Clear()
    {
        _spillAnchors.Clear();
        _occupiedBySpill.Clear();
    }
}
