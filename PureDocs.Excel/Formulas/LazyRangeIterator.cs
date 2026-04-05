namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Lazy iterator over worksheet cells in a range.
/// Skips empty/unused cells instead of materializing the full range.
/// Ideal for large range references like A:A (whole column).
/// </summary>
public sealed class LazyRangeIterator
{
    private readonly Worksheet _worksheet;
    private readonly int _startRow, _startCol, _endRow, _endCol;

    public int Rows => _endRow - _startRow + 1;
    public int Columns => _endCol - _startCol + 1;

    public LazyRangeIterator(Worksheet worksheet, string startRef, string endRef)
    {
        _worksheet = worksheet;
        CellReference.Parse(startRef, out _startRow, out _startCol);
        CellReference.Parse(endRef, out _endRow, out _endCol);
    }

    public LazyRangeIterator(Worksheet worksheet, int startRow, int startCol, int endRow, int endCol)
    {
        _worksheet = worksheet;
        _startRow = startRow;
        _startCol = startCol;
        _endRow = endRow;
        _endCol = endCol;
    }

    /// <summary>
    /// Iterates only over non-empty cells in the range.
    /// Returns (row, col, value) tuples.
    /// </summary>
    public IEnumerable<(int Row, int Column, FormulaValue Value)> EnumerateUsedCells()
    {
        var sheetData = _worksheet.GetSheetData();
        if (sheetData == null) yield break;

        foreach (var row in sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>())
        {
            var rowIdx = (int)(row.RowIndex?.Value ?? 0);
            if (rowIdx < _startRow || rowIdx > _endRow) continue;

            foreach (var cell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                var cellRef = cell.CellReference?.Value;
                if (cellRef == null) continue;

                CellReference.Parse(cellRef, out int r, out int c);
                if (c < _startCol || c > _endCol) continue;

                var wrapCell = _worksheet.GetCell(cellRef);
                var value = FormulaValue.FromObject(wrapCell.GetValue());
                yield return (r, c, value);
            }
        }
    }

    /// <summary>
    /// Counts non-empty cells without materializing a full array.
    /// </summary>
    public int CountUsedCells()
    {
        int count = 0;
        foreach (var _ in EnumerateUsedCells()) count++;
        return count;
    }

    /// <summary>
    /// Collects numeric values from non-empty cells (for SUM/AVERAGE on large ranges).
    /// </summary>
    public IEnumerable<double> EnumerateNumbers()
    {
        foreach (var (_, _, value) in EnumerateUsedCells())
        {
            if (value.IsNumber)
                yield return value.NumberValue;
        }
    }

    /// <summary>
    /// Materializes to ArrayValue only when needed (deferred).
    /// </summary>
    public ArrayValue ToArrayValue()
    {
        var arr = new ArrayValue(Rows, Columns);
        // Fill with blanks by default (ArrayValue initializes to default FormulaValue = Blank)

        foreach (var (row, col, value) in EnumerateUsedCells())
        {
            arr[row - _startRow, col - _startCol] = value;
        }

        return arr;
    }
}
