namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents a range of cells in a worksheet.
/// </summary>
public sealed class Range
{
    private readonly Worksheet _worksheet;
    private readonly string _rangeReference;
    private readonly int _startRow;
    private readonly int _startColumn;
    private readonly int _endRow;
    private readonly int _endColumn;

    internal Range(Worksheet worksheet, string rangeReference)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _rangeReference = rangeReference ?? throw new ArgumentNullException(nameof(rangeReference));

        ParseRangeReference(rangeReference, out _startRow, out _startColumn, out _endRow, out _endColumn);
    }

    /// <summary>
    /// Gets the range reference (e.g., "A1:B10").
    /// </summary>
    public string Address => _rangeReference;

    /// <summary>
    /// Gets the number of rows in the range.
    /// </summary>
    public int RowCount => _endRow - _startRow + 1;

    /// <summary>
    /// Gets the number of columns in the range.
    /// </summary>
    public int ColumnCount => _endColumn - _startColumn + 1;

    /// <summary>
    /// Sets values from a 2D array.
    /// </summary>
    public void SetValues(object[,] values)
    {
        if (values == null)
            throw new ArgumentNullException(nameof(values));

        int rows = values.GetLength(0);
        int cols = values.GetLength(1);

        for (int row = 0; row < rows && _startRow + row <= _endRow; row++)
        {
            for (int col = 0; col < cols && _startColumn + col <= _endColumn; col++)
            {
                var cell = _worksheet.GetCell(_startRow + row, _startColumn + col);
                var value = values[row, col];

                if (value is string str)
                    cell.SetValue(str);
                else if (value is double d)
                    cell.SetValue(d);
                else if (value is int i)
                    cell.SetValue(i);
                else if (value is bool b)
                    cell.SetValue(b);
                else if (value is DateTime dt)
                    cell.SetValue(dt);
                else if (value != null)
                    cell.SetValue(value.ToString()!);
            }
        }
    }

    /// <summary>
    /// Gets values as a 2D array.
    /// </summary>
    public object?[,] GetValues()
    {
        var result = new object?[RowCount, ColumnCount];

        for (int row = 0; row < RowCount; row++)
        {
            for (int col = 0; col < ColumnCount; col++)
            {
                var cell = _worksheet.GetCell(_startRow + row, _startColumn + col);
                result[row, col] = cell.GetValue();
            }
        }

        return result;
    }

    /// <summary>
    /// Auto-fits all columns in the range.
    /// </summary>
    public void AutoFit()
    {
        for (int col = _startColumn; col <= _endColumn; col++)
        {
            _worksheet.AutoFitColumn(col);
        }
    }

    /// <summary>
    /// Clears all cells in the range.
    /// </summary>
    public void Clear()
    {
        ForEachCell(cell => cell.Clear());
    }

    // ── Style Methods ──────────────────────────────────────────────────

    /// <summary>
    /// Applies a CellStyle to all cells in the range.
    /// </summary>
    public Range SetStyle(CellStyle style)
    {
        ForEachCell(cell => cell.Style = style);
        return this;
    }

    /// <summary>
    /// Applies a style using a fluent builder action to all cells in the range.
    /// Usage: range.ApplyStyle(s => s.SetBold().SetBackgroundColor(ExcelColor.Yellow));
    /// </summary>
    public Range ApplyStyle(Action<CellStyle> styleAction)
    {
        ForEachCell(cell => cell.ApplyStyle(styleAction));
        return this;
    }

    /// <summary>
    /// Sets all cells in the range to bold.
    /// </summary>
    public Range SetBold(bool bold = true)
    {
        ForEachCell(cell => cell.SetBold(bold));
        return this;
    }

    /// <summary>
    /// Sets the font size for all cells in the range.
    /// </summary>
    public Range SetFontSize(double size)
    {
        ForEachCell(cell => cell.SetFontSize(size));
        return this;
    }

    /// <summary>
    /// Sets the font color for all cells in the range.
    /// </summary>
    public Range SetFontColor(ExcelColor color)
    {
        ForEachCell(cell => cell.SetFontColor(color));
        return this;
    }

    /// <summary>
    /// Sets the font name for all cells in the range.
    /// </summary>
    public Range SetFontName(string name)
    {
        ForEachCell(cell => cell.SetFontName(name));
        return this;
    }

    /// <summary>
    /// Sets the background color for all cells in the range.
    /// </summary>
    public Range SetBackgroundColor(ExcelColor color)
    {
        ForEachCell(cell => cell.SetBackgroundColor(color));
        return this;
    }

    /// <summary>
    /// Sets the background color from hex string for all cells in the range.
    /// </summary>
    public Range SetBackgroundColor(string hex)
    {
        var color = ExcelColor.FromHex(hex);
        ForEachCell(cell => cell.SetBackgroundColor(color));
        return this;
    }

    /// <summary>
    /// Sets all borders for all cells in the range.
    /// </summary>
    public Range SetAllBorders(ExcelBorderStyle style, ExcelColor? color = null)
    {
        ForEachCell(cell => cell.SetAllBorders(style, color));
        return this;
    }

    /// <summary>
    /// Sets the horizontal alignment for all cells in the range.
    /// </summary>
    public Range SetHorizontalAlignment(ExcelHorizontalAlignment alignment)
    {
        ForEachCell(cell => cell.SetHorizontalAlignment(alignment));
        return this;
    }

    /// <summary>
    /// Sets the vertical alignment for all cells in the range.
    /// </summary>
    public Range SetVerticalAlignment(ExcelVerticalAlignment alignment)
    {
        ForEachCell(cell => cell.SetVerticalAlignment(alignment));
        return this;
    }

    /// <summary>
    /// Sets text wrapping for all cells in the range.
    /// </summary>
    public Range SetWrapText(bool wrap = true)
    {
        ForEachCell(cell => cell.SetWrapText(wrap));
        return this;
    }

    /// <summary>
    /// Sets the number format for all cells in the range.
    /// </summary>
    public Range SetNumberFormat(ExcelNumberFormat format)
    {
        ForEachCell(cell => cell.SetNumberFormat(format));
        return this;
    }

    /// <summary>
    /// Sets a custom number format for all cells in the range.
    /// </summary>
    public Range SetNumberFormat(string formatCode)
    {
        ForEachCell(cell => cell.SetNumberFormat(formatCode));
        return this;
    }

    // ── Helper ─────────────────────────────────────────────────────────

    /// <summary>
    /// Executes an action on each cell in the range.
    /// </summary>
    private void ForEachCell(Action<Cell> action)
    {
        for (int row = _startRow; row <= _endRow; row++)
        {
            for (int col = _startColumn; col <= _endColumn; col++)
            {
                action(_worksheet.GetCell(row, col));
            }
        }
    }

    private static void ParseRangeReference(string rangeRef, out int startRow, out int startCol, out int endRow, out int endCol)
    {
        var parts = rangeRef.Split(':');
        
        if (parts.Length != 2)
            throw new ArgumentException($"Invalid range reference: {rangeRef}", nameof(rangeRef));

        CellReference.Parse(parts[0], out startRow, out startCol);
        CellReference.Parse(parts[1], out endRow, out endCol);

        if (startRow > endRow || startCol > endCol)
            throw new ArgumentException($"Invalid range reference: {rangeRef}", nameof(rangeRef));
    }
}
