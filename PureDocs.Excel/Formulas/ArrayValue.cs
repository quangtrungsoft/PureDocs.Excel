namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// 2D array value for formula results and range data.
/// Row-major storage for cache-friendly access.
/// </summary>
/// <remarks>
/// <para>
/// IMPORTANT: The indexers in this class return <see cref="FormulaValue.ErrorRef"/> for out-of-bounds
/// access instead of throwing <see cref="IndexOutOfRangeException"/>. This is intentional to match
/// Excel's behavior where invalid array references result in #REF! errors rather than crashes.
/// </para>
/// <para>
/// For safer access patterns, use <see cref="TryGet(int, int, out FormulaValue)"/> which returns
/// a boolean indicating success.
/// </para>
/// </remarks>
public sealed class ArrayValue
{
    public static readonly ArrayValue Empty = new(0, 0);

    private readonly FormulaValue[] _data;
    public int Rows { get; }
    public int Columns { get; }
    public int Length => _data.Length;

    public ArrayValue(int rows, int columns)
    {
        Rows = rows;
        Columns = columns;
        _data = new FormulaValue[rows * columns];
    }

    /// <summary>
    /// Gets or sets the value at the specified row and column.
    /// </summary>
    /// <remarks>
    /// Returns <see cref="FormulaValue.ErrorRef"/> for out-of-bounds access (Excel behavior).
    /// Use <see cref="TryGet(int, int, out FormulaValue)"/> for explicit bounds checking.
    /// </remarks>
    public FormulaValue this[int row, int col]
    {
        get => (row >= 0 && row < Rows && col >= 0 && col < Columns)
            ? _data[row * Columns + col] : FormulaValue.ErrorRef;
        set { if (row >= 0 && row < Rows && col >= 0 && col < Columns) _data[row * Columns + col] = value; }
    }

    /// <summary>
    /// Gets or sets the value at the specified flat index.
    /// </summary>
    /// <remarks>
    /// Returns <see cref="FormulaValue.ErrorRef"/> for out-of-bounds access (Excel behavior).
    /// Use <see cref="TryGet(int, out FormulaValue)"/> for explicit bounds checking.
    /// </remarks>
    public FormulaValue this[int index]
    {
        get => (index >= 0 && index < _data.Length) ? _data[index] : FormulaValue.ErrorRef;
        set { if (index >= 0 && index < _data.Length) _data[index] = value; }
    }

    /// <summary>
    /// Attempts to get a value at the specified row and column.
    /// </summary>
    /// <returns>True if the indices are valid; false otherwise.</returns>
    public bool TryGet(int row, int col, out FormulaValue value)
    {
        if (row >= 0 && row < Rows && col >= 0 && col < Columns)
        {
            value = _data[row * Columns + col];
            return true;
        }
        value = FormulaValue.ErrorRef;
        return false;
    }

    /// <summary>
    /// Attempts to get a value at the specified flat index.
    /// </summary>
    /// <returns>True if the index is valid; false otherwise.</returns>
    public bool TryGet(int index, out FormulaValue value)
    {
        if (index >= 0 && index < _data.Length)
        {
            value = _data[index];
            return true;
        }
        value = FormulaValue.ErrorRef;
        return false;
    }

    /// <summary>
    /// Checks if the specified row and column are within bounds.
    /// </summary>
    public bool IsInBounds(int row, int col) => row >= 0 && row < Rows && col >= 0 && col < Columns;

    /// <summary>Iterate all values (row-major order).</summary>
    public ReadOnlySpan<FormulaValue> Values => _data.AsSpan();

    /// <summary>Iterate values as IEnumerable (for LINQ).</summary>
    public IEnumerable<FormulaValue> EnumerateValues() => _data;

    /// <summary>Convert to object?[] for backward compatibility.</summary>
    public object?[] ToObjectArray()
    {
        var result = new object?[_data.Length];
        for (int i = 0; i < _data.Length; i++)
            result[i] = _data[i].ToObject();
        return result;
    }

    /// <summary>Create from flat FormulaValue array (single column).</summary>
    public static ArrayValue FromFlat(FormulaValue[] values)
    {
        var arr = new ArrayValue(values.Length, 1);
        for (int i = 0; i < values.Length; i++)
            arr._data[i] = values[i];
        return arr;
    }

    /// <summary>Create from flat list (single column).</summary>
    public static ArrayValue FromList(List<FormulaValue> values)
    {
        var arr = new ArrayValue(values.Count, 1);
        for (int i = 0; i < values.Count; i++)
            arr._data[i] = values[i];
        return arr;
    }

    /// <summary>Create a 1×1 array from a scalar value for broadcasting.</summary>
    public static ArrayValue FromScalar(FormulaValue value)
    {
        var arr = new ArrayValue(1, 1);
        arr._data[0] = value;
        return arr;
    }

    // ── Array Broadcasting ────────────────────────────────────────────

    /// <summary>
    /// Excel array broadcasting: apply binary operation element-wise.
    /// Dimensions are broadcast: (1×N) op (M×1) → (M×N).
    /// Scalars are broadcast to any dimension.
    /// </summary>
    public static ArrayValue Broadcast(ArrayValue a, ArrayValue b,
        Func<FormulaValue, FormulaValue, FormulaValue> op)
    {
        int rows = Math.Max(a.Rows, b.Rows);
        int cols = Math.Max(a.Columns, b.Columns);

        // Validate broadcasting compatibility
        if ((a.Rows != 1 && b.Rows != 1 && a.Rows != b.Rows) ||
            (a.Columns != 1 && b.Columns != 1 && a.Columns != b.Columns))
        {
            // Incompatible dimensions → #VALUE! array
            var errArr = new ArrayValue(rows, cols);
            for (int i = 0; i < errArr._data.Length; i++)
                errArr._data[i] = FormulaValue.ErrorValue;
            return errArr;
        }

        var result = new ArrayValue(rows, cols);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                var va = a[r < a.Rows ? r : 0, c < a.Columns ? c : 0];
                var vb = b[r < b.Rows ? r : 0, c < b.Columns ? c : 0];
                result[r, c] = op(va, vb);
            }
        return result;
    }

    /// <summary>
    /// Apply unary operation element-wise to all values in the array.
    /// </summary>
    public static ArrayValue Map(ArrayValue a, Func<FormulaValue, FormulaValue> op)
    {
        var result = new ArrayValue(a.Rows, a.Columns);
        for (int i = 0; i < a._data.Length; i++)
            result._data[i] = op(a._data[i]);
        return result;
    }

    // ── Implicit Intersection ─────────────────────────────────────────

    /// <summary>
    /// Implicit intersection: given a formula's row, extract the relevant value.
    /// For a single-column array → return value at formulaRow.
    /// For a single-row array → return value at formulaCol.
    /// For 1×1 → return that value. Otherwise → #VALUE!.
    /// </summary>
    public FormulaValue ImplicitIntersect(int formulaRow, int formulaCol,
        int rangeStartRow, int rangeStartCol)
    {
        if (Rows == 1 && Columns == 1)
            return this[0, 0];

        if (Columns == 1)
        {
            // Single column: intersect with formula's row
            int offset = formulaRow - rangeStartRow;
            if (offset >= 0 && offset < Rows)
                return this[offset, 0];
            return FormulaValue.ErrorValue;
        }

        if (Rows == 1)
        {
            // Single row: intersect with formula's column
            int offset = formulaCol - rangeStartCol;
            if (offset >= 0 && offset < Columns)
                return this[0, offset];
            return FormulaValue.ErrorValue;
        }

        // Multi-row AND multi-column: can intersect if both row and col match
        int rowOffset = formulaRow - rangeStartRow;
        int colOffset = formulaCol - rangeStartCol;
        if (rowOffset >= 0 && rowOffset < Rows && colOffset >= 0 && colOffset < Columns)
            return this[rowOffset, colOffset];

        return FormulaValue.ErrorValue;
    }

    // ── Row/Column extraction ─────────────────────────────────────────

    /// <summary>Extract a single row as a new 1×N array.</summary>
    public ArrayValue GetRow(int row)
    {
        if (row < 0 || row >= Rows) return Empty;
        var result = new ArrayValue(1, Columns);
        for (int c = 0; c < Columns; c++)
            result._data[c] = this[row, c];
        return result;
    }

    /// <summary>Extract a single column as a new M×1 array.</summary>
    public ArrayValue GetColumn(int col)
    {
        if (col < 0 || col >= Columns) return Empty;
        var result = new ArrayValue(Rows, 1);
        for (int r = 0; r < Rows; r++)
            result._data[r] = this[r, col];
        return result;
    }

    /// <summary>Transpose the array (swap rows and columns).</summary>
    public ArrayValue Transpose()
    {
        var result = new ArrayValue(Columns, Rows);
        for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Columns; c++)
                result[c, r] = this[r, c];
        return result;
    }

    /// <summary>Create a 1×1 array with a single value.</summary>
    public static ArrayValue Scalar(FormulaValue value)
    {
        var arr = new ArrayValue(1, 1);
        arr._data[0] = value;
        return arr;
    }
}
