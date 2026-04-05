namespace TVE.PureDocs.Excel.Formulas;

internal static class LookupFunctions
{
    public static void Register(FunctionRegistry r)
    {
        r.Register("VLOOKUP", VLookup, 3, 4); r.Register("HLOOKUP", HLookup, 3, 4);
        r.Register("INDEX", Index, 2, 3); r.Register("MATCH", Match, 2, 3);
        r.Register("CHOOSE", Choose, 2); r.Register("ROW", Row, 0, 1);
        r.Register("COLUMN", Column, 0, 1); r.Register("ROWS", Rows, 1, 1);
        r.Register("COLUMNS", Columns, 1, 1);
        r.Register("OFFSET", Offset, 3, 5, isVolatile: true);
    }

    /// <summary>
    /// VLOOKUP - Vertical lookup in a table.
    /// </summary>
    /// <remarks>
    /// IMPORTANT: When using approximate match (range_lookup = TRUE or omitted),
    /// the first column of the table MUST be sorted in ascending order.
    /// If the data is not sorted, this function may return incorrect results.
    /// This matches Excel's documented behavior.
    /// </remarks>
    private static FormulaValue VLookup(List<FormulaNode> a, FormulaContext c)
    {
        var lookupVal = a[0].Evaluate(c);
        if (lookupVal.IsError) return lookupVal;
        var tableVal = a[1].Evaluate(c);
        if (tableVal.IsError) return tableVal;
        if (!tableVal.IsArray) return FormulaValue.ErrorRef;
        if (!FormulaHelper.TryEvalDouble(a[2], c, out double ci, out var e)) return e;
        int colIdx = (int)ci - 1;
        bool exactMatch = true;
        if (a.Count > 3) { var mv = a[3].Evaluate(c); if (mv.IsError) return mv; exactMatch = mv.CoerceToBool().BooleanValue == false; }

        var tbl = tableVal.ArrayVal;
        if (colIdx < 0 || colIdx >= tbl.Columns) return FormulaValue.ErrorRef;

        if (exactMatch)
        {
            for (int r = 0; r < tbl.Rows; r++)
                if (FormulaValue.AreEqual(tbl[r, 0], lookupVal)) return tbl[r, colIdx];
            return FormulaValue.ErrorNA;
        }
        // Approximate match: find largest value <= lookup
        // WARNING: This assumes the first column is sorted in ascending order.
        // If the data is not sorted, results will be incorrect.
        int best = -1;
        for (int r = 0; r < tbl.Rows; r++)
        {
            if (FormulaValue.Compare(tbl[r, 0], lookupVal) <= 0)
                best = r;
            else break; // Stop at first value > lookup (assumes sorted data)
        }
        return best >= 0 ? tbl[best, colIdx] : FormulaValue.ErrorNA;
    }

    /// <summary>
    /// HLOOKUP - Horizontal lookup in a table.
    /// </summary>
    /// <remarks>
    /// IMPORTANT: When using approximate match (range_lookup = TRUE or omitted),
    /// the first row of the table MUST be sorted in ascending order.
    /// If the data is not sorted, this function may return incorrect results.
    /// </remarks>
    private static FormulaValue HLookup(List<FormulaNode> a, FormulaContext c)
    {
        var lookupVal = a[0].Evaluate(c);
        if (lookupVal.IsError) return lookupVal;
        var tableVal = a[1].Evaluate(c);
        if (tableVal.IsError) return tableVal;
        if (!tableVal.IsArray) return FormulaValue.ErrorRef;
        if (!FormulaHelper.TryEvalDouble(a[2], c, out double ri, out var e)) return e;
        int rowIdx = (int)ri - 1;
        bool exactMatch = true;
        if (a.Count > 3) { var mv = a[3].Evaluate(c); if (mv.IsError) return mv; exactMatch = mv.CoerceToBool().BooleanValue == false; }

        var tbl = tableVal.ArrayVal;
        if (rowIdx < 0 || rowIdx >= tbl.Rows) return FormulaValue.ErrorRef;

        if (exactMatch)
        {
            for (int col = 0; col < tbl.Columns; col++)
                if (FormulaValue.AreEqual(tbl[0, col], lookupVal)) return tbl[rowIdx, col];
            return FormulaValue.ErrorNA;
        }
        // Approximate match: assumes first row is sorted ascending
        int best = -1;
        for (int col = 0; col < tbl.Columns; col++)
        {
            if (FormulaValue.Compare(tbl[0, col], lookupVal) <= 0) best = col; else break;
        }
        return best >= 0 ? tbl[rowIdx, best] : FormulaValue.ErrorNA;
    }

    private static FormulaValue Index(List<FormulaNode> a, FormulaContext c)
    {
        var arrVal = a[0].Evaluate(c);
        if (arrVal.IsError) return arrVal;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double rv, out var e)) return e;
        int row = (int)rv - 1;
        int col = 0;
        if (a.Count > 2) { if (!FormulaHelper.TryEvalDouble(a[2], c, out double cv, out e)) return e; col = (int)cv - 1; }

        if (!arrVal.IsArray)
            return row == 0 && col == 0 ? arrVal : FormulaValue.ErrorRef;

        var arr = arrVal.ArrayVal;
        // INDEX with row=0 returns entire column as array
        if (row < 0 && col >= 0 && col < arr.Columns)
        {
            var result = new ArrayValue(arr.Rows, 1);
            for (int r = 0; r < arr.Rows; r++) result[r, 0] = arr[r, col];
            return FormulaValue.Array(result);
        }
        // INDEX with col=0 returns entire row as array
        if (col < 0 && row >= 0 && row < arr.Rows)
        {
            var result = new ArrayValue(1, arr.Columns);
            for (int cc = 0; cc < arr.Columns; cc++) result[0, cc] = arr[row, cc];
            return FormulaValue.Array(result);
        }

        if (row < 0 || row >= arr.Rows || col < 0 || col >= arr.Columns)
            return FormulaValue.ErrorRef;
        return arr[row, col];
    }

    private static FormulaValue Match(List<FormulaNode> a, FormulaContext c)
    {
        var lookupVal = a[0].Evaluate(c);
        if (lookupVal.IsError) return lookupVal;
        var arrVal = a[1].Evaluate(c);
        if (arrVal.IsError) return arrVal;
        int matchType = 1;
        if (a.Count > 2) { if (!FormulaHelper.TryEvalDouble(a[2], c, out double mt, out var e)) return e; matchType = (int)mt; }

        if (!arrVal.IsArray) return FormulaValue.AreEqual(arrVal, lookupVal) ? FormulaValue.One : FormulaValue.ErrorNA;

        var arr = arrVal.ArrayVal;
        int len = arr.Length;

        if (matchType == 0)
        {
            for (int i = 0; i < len; i++)
                if (FormulaValue.AreEqual(arr[i], lookupVal)) return FormulaValue.Number(i + 1);
            return FormulaValue.ErrorNA;
        }
        if (matchType == 1)
        {
            int best = -1;
            for (int i = 0; i < len; i++)
                if (FormulaValue.Compare(arr[i], lookupVal) <= 0) best = i; else break;
            return best >= 0 ? FormulaValue.Number(best + 1) : FormulaValue.ErrorNA;
        }
        // matchType == -1
        int bestR = -1;
        for (int i = 0; i < len; i++)
            if (FormulaValue.Compare(arr[i], lookupVal) >= 0) bestR = i; else break;
        return bestR >= 0 ? FormulaValue.Number(bestR + 1) : FormulaValue.ErrorNA;
    }

    private static FormulaValue Choose(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double iv, out var e)) return e;
        int idx = (int)iv;
        if (idx < 1 || idx >= a.Count) return FormulaValue.ErrorValue;
        return a[idx].Evaluate(c);
    }

    private static FormulaValue Row(List<FormulaNode> a, FormulaContext c)
    {
        if (a.Count == 0) return FormulaValue.Number(1);
        if (a[0] is CellReferenceNode cr)
        {
            CellReference.Parse(cr.Reference, out int r, out _);
            return FormulaValue.Number(r);
        }
        if (a[0] is RangeReferenceNode rr)
        {
            CellReference.Parse(rr.StartRef, out int r, out _);
            return FormulaValue.Number(r);
        }
        return FormulaValue.ErrorValue;
    }

    private static FormulaValue Column(List<FormulaNode> a, FormulaContext c)
    {
        if (a.Count == 0) return FormulaValue.Number(1);
        if (a[0] is CellReferenceNode cr)
        {
            CellReference.Parse(cr.Reference, out _, out int col);
            return FormulaValue.Number(col);
        }
        if (a[0] is RangeReferenceNode rr)
        {
            CellReference.Parse(rr.StartRef, out _, out int col);
            return FormulaValue.Number(col);
        }
        return FormulaValue.ErrorValue;
    }

    private static FormulaValue Rows(List<FormulaNode> a, FormulaContext c)
    {
        if (a[0] is RangeReferenceNode rr)
        {
            var bounds = c.GetRangeBounds(rr.StartRef, rr.EndRef);
            return FormulaValue.Number(bounds.endRow - bounds.startRow + 1);
        }
        var v = a[0].Evaluate(c);
        return v.IsArray ? FormulaValue.Number(v.ArrayVal.Rows) : FormulaValue.One;
    }

    private static FormulaValue Columns(List<FormulaNode> a, FormulaContext c)
    {
        if (a[0] is RangeReferenceNode rr)
        {
            var bounds = c.GetRangeBounds(rr.StartRef, rr.EndRef);
            return FormulaValue.Number(bounds.endCol - bounds.startCol + 1);
        }
        var v = a[0].Evaluate(c);
        return v.IsArray ? FormulaValue.Number(v.ArrayVal.Columns) : FormulaValue.One;
    }

    private static FormulaValue Offset(List<FormulaNode> a, FormulaContext c)
    {
        if (a[0] is not CellReferenceNode cr && a[0] is not RangeReferenceNode)
            return FormulaValue.ErrorValue;

        string baseRef = a[0] is CellReferenceNode cref ? cref.Reference
            : ((RangeReferenceNode)a[0]).StartRef;

        CellReference.Parse(baseRef, out int baseRow, out int baseCol);
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double rowOff, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[2], c, out double colOff, out e)) return e;

        int newRow = baseRow + (int)rowOff;
        int newCol = baseCol + (int)colOff;
        if (newRow < 1 || newCol < 1) return FormulaValue.ErrorRef;

        int height = 1, width = 1;
        if (a.Count > 3) { if (!FormulaHelper.TryEvalDouble(a[3], c, out double h, out e)) return e; height = (int)h; }
        if (a.Count > 4) { if (!FormulaHelper.TryEvalDouble(a[4], c, out double w, out e)) return e; width = (int)w; }

        if (height == 1 && width == 1)
            return c.GetCellValue(CellReference.FromRowColumn(newRow, newCol));

        string start = CellReference.FromRowColumn(newRow, newCol);
        string end = CellReference.FromRowColumn(newRow + height - 1, newCol + width - 1);
        return c.GetRangeValues(start, end);
    }
}
