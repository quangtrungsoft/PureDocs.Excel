namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Represents a rectangular area of cells (e.g., A1:D10) as a compact struct.
/// Used instead of expanding ranges into individual CellAddress entries.
/// Memory: 16 bytes per area vs 20 bytes × N cells.
/// </summary>
public readonly struct RangeArea : IEquatable<RangeArea>
{
    public readonly int StartRow;
    public readonly int StartCol;
    public readonly int EndRow;
    public readonly int EndCol;

    public RangeArea(int startRow, int startCol, int endRow, int endCol)
    {
        StartRow = Math.Min(startRow, endRow);
        StartCol = Math.Min(startCol, endCol);
        EndRow = Math.Max(startRow, endRow);
        EndCol = Math.Max(startCol, endCol);
    }

    /// <summary>Creates from cell reference strings.</summary>
    public static RangeArea FromReferences(string startRef, string endRef)
    {
        CellReference.Parse(startRef, out int sr, out int sc);
        CellReference.Parse(endRef, out int er, out int ec);
        return new RangeArea(sr, sc, er, ec);
    }

    /// <summary>
    /// Total number of cells in this area.
    /// Uses long to handle full column/row ranges (max ~17 billion cells).
    /// </summary>
    /// <remarks>
    /// Full sheet range A1:XFD1048576 = 1,048,576 rows × 16,384 cols = 17,179,869,184 cells.
    /// This exceeds int32 max (2.1 billion), so long is required.
    /// </remarks>
    public long CellCount => (long)(EndRow - StartRow + 1) * (EndCol - StartCol + 1);

    /// <summary>Checks if a cell address falls within this area.</summary>
    public bool Contains(CellAddress cell)
        => cell.Row >= StartRow && cell.Row <= EndRow
        && cell.Column >= StartCol && cell.Column <= EndCol;

    /// <summary>Checks if a cell (row, col) falls within this area.</summary>
    public bool Contains(int row, int col)
        => row >= StartRow && row <= EndRow
        && col >= StartCol && col <= EndCol;

    /// <summary>Checks if this area overlaps with another area.</summary>
    public bool Overlaps(RangeArea other)
        => StartRow <= other.EndRow && EndRow >= other.StartRow
        && StartCol <= other.EndCol && EndCol >= other.StartCol;

    public bool Equals(RangeArea other)
        => StartRow == other.StartRow && StartCol == other.StartCol
        && EndRow == other.EndRow && EndCol == other.EndCol;
    public override bool Equals(object? obj) => obj is RangeArea o && Equals(o);
    public override int GetHashCode() => HashCode.Combine(StartRow, StartCol, EndRow, EndCol);
    public override string ToString()
        => $"{CellReference.FromRowColumn(StartRow, StartCol)}:{CellReference.FromRowColumn(EndRow, EndCol)}";
}

/// <summary>
/// Compact set of cell dependencies — stores individual cells AND range areas.
/// 
/// Key optimization: SUM(A:A) becomes 1 RangeArea (16 bytes)
/// instead of 1,048,576 CellAddress entries (~20MB).
/// 
/// Query "does cell X affect this formula?" is O(singles + areas).
/// For typical formulas (1-5 refs + 0-2 ranges), this is effectively O(1).
/// </summary>
public sealed class RangeSet
{
    private readonly HashSet<CellAddress> _singles = new();
    private readonly List<RangeArea> _areas = new();

    /// <summary>Number of individual cell references.</summary>
    public int SingleCount => _singles.Count;

    /// <summary>Number of range area references.</summary>
    public int AreaCount => _areas.Count;

    /// <summary>Adds a single cell reference.</summary>
    public void AddCell(CellAddress cell) => _singles.Add(cell);

    /// <summary>Adds a range area reference (compact, no cell expansion).</summary>
    public void AddRange(RangeArea area) => _areas.Add(area);

    /// <summary>Adds a range from row/col bounds.</summary>
    public void AddRange(int startRow, int startCol, int endRow, int endCol)
        => _areas.Add(new RangeArea(startRow, startCol, endRow, endCol));

    /// <summary>
    /// Checks if a cell is contained in this dependency set.
    /// O(1) for singles + O(areas) for ranges.
    /// </summary>
    public bool Contains(CellAddress cell)
    {
        if (_singles.Contains(cell)) return true;
        for (int i = 0; i < _areas.Count; i++)
            if (_areas[i].Contains(cell)) return true;
        return false;
    }

    /// <summary>Checks if any area overlaps with given area.</summary>
    public bool OverlapsAny(RangeArea area)
    {
        for (int i = 0; i < _areas.Count; i++)
            if (_areas[i].Overlaps(area)) return true;
        // Also check if any single cell falls within the area
        foreach (var cell in _singles)
            if (area.Contains(cell)) return true;
        return false;
    }

    /// <summary>Returns all individual cells (for backward compat). Warning: expensive for large ranges.</summary>
    public HashSet<CellAddress> ExpandAll()
    {
        var result = new HashSet<CellAddress>(_singles);
        foreach (var area in _areas)
            for (int r = area.StartRow; r <= area.EndRow; r++)
                for (int c = area.StartCol; c <= area.EndCol; c++)
                    result.Add(new CellAddress(r, c));
        return result;
    }

    /// <summary>Gets all range areas.</summary>
    public IReadOnlyList<RangeArea> Areas => _areas;

    /// <summary>Gets all single cells.</summary>
    public IReadOnlySet<CellAddress> Singles => _singles;

    /// <summary>Estimates memory usage in bytes.</summary>
    public long EstimatedMemoryBytes
        => (_singles.Count * 20L) + (_areas.Count * 16L) + 64L; // overhead
}
