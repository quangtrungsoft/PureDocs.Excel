namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Lightweight cell address for dependency tracking.
/// </summary>
public readonly struct CellAddress : IEquatable<CellAddress>
{
    public readonly int SheetIndex; // -1 = current sheet
    public readonly int Row;
    public readonly int Column;

    public CellAddress(int row, int column, int sheetIndex = -1)
    {
        Row = row;
        Column = column;
        SheetIndex = sheetIndex;
    }

    /// <summary>Parse from reference string like "A1", "$B$2".</summary>
    public static CellAddress FromReference(string cellRef, int sheetIndex = -1)
    {
        CellReference.Parse(cellRef, out int row, out int col);
        return new CellAddress(row, col, sheetIndex);
    }

    /// <summary>Convert back to reference string.</summary>
    public string ToReference() => CellReference.FromRowColumn(Row, Column);

    public bool Equals(CellAddress other) => Row == other.Row && Column == other.Column && SheetIndex == other.SheetIndex;
    public override bool Equals(object? obj) => obj is CellAddress o && Equals(o);
    public override int GetHashCode() => HashCode.Combine(SheetIndex, Row, Column);
    public override string ToString() => SheetIndex >= 0 ? $"Sheet{SheetIndex}!{ToReference()}" : ToReference();
    public static bool operator ==(CellAddress l, CellAddress r) => l.Equals(r);
    public static bool operator !=(CellAddress l, CellAddress r) => !l.Equals(r);
}
