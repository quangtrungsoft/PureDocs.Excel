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
        r.Register("LOOKUP", Lookup, 2, 3);
        r.Register("XLOOKUP", XLookup, 3, 6);
        r.Register("ADDRESS", Address, 2, 5);
        r.Register("SUBTOTAL", Subtotal, 2);
        // Volatile: the reference is computed at runtime, so its precedents cannot be
        // tracked statically — mark volatile so recalc always re-evaluates it.
        r.Register("INDIRECT", Indirect, 1, 2, isVolatile: true);
    }

    // ── INDIRECT ─────────────────────────────────────────────────────

    /// <summary>
    /// INDIRECT(ref_text, [a1]) — interprets a text string as a reference and returns its value(s).
    /// Only A1-style references are supported; a1 = FALSE (R1C1) yields #REF!.
    /// </summary>
    private static FormulaValue Indirect(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalString(a[0], c, out string refText, out var e)) return e;
        if (string.IsNullOrWhiteSpace(refText)) return FormulaValue.ErrorRef;

        bool a1 = true;
        if (a.Count > 1) { var v = a[1].Evaluate(c); if (v.IsError) return v; a1 = v.CoerceToBool().BooleanValue; }
        if (!a1) return FormulaValue.ErrorRef; // R1C1 style not supported

        // Optional sheet prefix: Sheet2!A1  or  'My Sheet'!A1:B5
        string cellPart = refText;
        string? sheetName = null;
        int bang = refText.IndexOf('!');
        if (bang > 0)
        {
            sheetName = refText[..bang].Trim('\'');
            cellPart = refText[(bang + 1)..];
        }

        try
        {
            int colon = cellPart.IndexOf(':');
            if (colon > 0)
            {
                string start = cellPart[..colon], end = cellPart[(colon + 1)..];
                CellReference.Parse(start, out _, out _);   // validate (throws on bad ref)
                CellReference.Parse(end, out _, out _);
                return sheetName != null
                    ? c.GetSheetRangeValues(sheetName, start, end)
                    : c.GetRangeValues(start, end);
            }

            CellReference.Parse(cellPart, out _, out _);     // validate
            return sheetName != null
                ? c.GetSheetCellValue(sheetName, cellPart)
                : c.GetCellValue(cellPart);
        }
        catch
        {
            return FormulaValue.ErrorRef;
        }
    }

    // ── LOOKUP (vector form) ─────────────────────────────────────────

    /// <summary>LOOKUP(lookup_value, lookup_vector, [result_vector]) — approximate match, ascending.</summary>
    private static FormulaValue Lookup(List<FormulaNode> a, FormulaContext c)
    {
        var lookupVal = a[0].Evaluate(c);
        if (lookupVal.IsError) return lookupVal;
        var vecVal = a[1].Evaluate(c);
        if (vecVal.IsError) return vecVal;
        var vec = vecVal.IsArray ? vecVal.ArrayVal : WrapSingle(vecVal);

        ArrayValue? res = null;
        if (a.Count > 2)
        {
            var rv = a[2].Evaluate(c);
            if (rv.IsError) return rv;
            res = rv.IsArray ? rv.ArrayVal : WrapSingle(rv);
        }

        // Largest value <= lookup (assumes ascending order, matching Excel).
        int best = -1;
        for (int i = 0; i < vec.Length; i++)
        {
            if (FormulaValue.Compare(vec[i], lookupVal) <= 0) best = i;
            else break;
        }
        if (best < 0) return FormulaValue.ErrorNA;
        if (res != null) return best < res.Length ? res[best] : FormulaValue.ErrorNA;
        return vec[best];
    }

    // ── XLOOKUP ──────────────────────────────────────────────────────

    /// <summary>
    /// XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode]).
    /// match_mode: 0 exact (default), -1 exact/next-smaller, 1 exact/next-larger, 2 wildcard.
    /// search_mode: 1 first-to-last (default), -1 last-to-first.
    /// </summary>
    private static FormulaValue XLookup(List<FormulaNode> a, FormulaContext c)
    {
        var lookupVal = a[0].Evaluate(c);
        if (lookupVal.IsError) return lookupVal;
        var searchVal = a[1].Evaluate(c);
        if (searchVal.IsError) return searchVal;
        var returnVal = a[2].Evaluate(c);
        if (returnVal.IsError) return returnVal;

        var search = searchVal.IsArray ? searchVal.ArrayVal : WrapSingle(searchVal);
        var ret = returnVal.IsArray ? returnVal.ArrayVal : WrapSingle(returnVal);

        FormulaValue? ifNotFound = null;
        if (a.Count > 3) { var v = a[3].Evaluate(c); if (v.IsError) return v; ifNotFound = v; }
        int matchMode = 0;
        if (a.Count > 4) { if (!FormulaHelper.TryEvalDouble(a[4], c, out double mm, out var e)) return e; matchMode = (int)mm; }
        int searchMode = 1;
        if (a.Count > 5) { if (!FormulaHelper.TryEvalDouble(a[5], c, out double sm, out var e)) return e; searchMode = (int)sm; }

        int n = search.Length;
        int found = -1;

        if (matchMode == 0 || matchMode == 2)
        {
            int start = searchMode == -1 ? n - 1 : 0;
            int step = searchMode == -1 ? -1 : 1;
            for (int i = start; i >= 0 && i < n; i += step)
            {
                bool m = matchMode == 2
                    ? FormulaHelper.MatchesCriteria(search[i], lookupVal.AsText())
                    : FormulaValue.AreEqual(search[i], lookupVal);
                if (m) { found = i; break; }
            }
        }
        else
        {
            int bestSmaller = -1, bestLarger = -1;
            for (int i = 0; i < n; i++)
            {
                int cmp = FormulaValue.Compare(search[i], lookupVal);
                if (cmp == 0) { found = i; break; }
                if (cmp < 0) { if (bestSmaller < 0 || FormulaValue.Compare(search[i], search[bestSmaller]) > 0) bestSmaller = i; }
                else { if (bestLarger < 0 || FormulaValue.Compare(search[i], search[bestLarger]) < 0) bestLarger = i; }
            }
            if (found < 0) found = matchMode == -1 ? bestSmaller : bestLarger;
        }

        if (found < 0)
            return ifNotFound ?? FormulaValue.ErrorNA;

        // Vertical search whose return array spans multiple columns → return the matched row.
        if (ret.Rows == n && ret.Columns > 1)
        {
            var row = new ArrayValue(1, ret.Columns);
            for (int cc = 0; cc < ret.Columns; cc++) row[0, cc] = ret[found, cc];
            return FormulaValue.Array(row);
        }
        return found < ret.Length ? ret[found] : FormulaValue.ErrorNA;
    }

    // ── ADDRESS ──────────────────────────────────────────────────────

    /// <summary>ADDRESS(row, column, [abs_num], [a1], [sheet_text]) — builds a reference as text.</summary>
    private static FormulaValue Address(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double rv, out var e)) return e;
        if (!FormulaHelper.TryEvalDouble(a[1], c, out double cv, out e)) return e;
        int row = (int)rv, col = (int)cv;
        if (row < 1 || col < 1) return FormulaValue.ErrorValue;

        int absNum = 1;
        if (a.Count > 2) { if (!FormulaHelper.TryEvalDouble(a[2], c, out double an, out e)) return e; absNum = (int)an; }
        if (absNum < 1 || absNum > 4) return FormulaValue.ErrorValue;
        bool a1 = true;
        if (a.Count > 3) { var v = a[3].Evaluate(c); if (v.IsError) return v; a1 = v.CoerceToBool().BooleanValue; }
        string? sheet = null;
        if (a.Count > 4) { if (!FormulaHelper.TryEvalString(a[4], c, out sheet, out e)) return e; }

        string reference;
        if (a1)
        {
            string colAbs = (absNum == 1 || absNum == 3) ? "$" : "";
            string rowAbs = (absNum == 1 || absNum == 2) ? "$" : "";
            reference = $"{colAbs}{ColumnLetter(col)}{rowAbs}{row}";
        }
        else
        {
            // R1C1: 1 absolute, 2 abs-row/rel-col, 3 rel-row/abs-col, 4 relative.
            string rPart = (absNum == 1 || absNum == 2) ? $"R{row}" : $"R[{row}]";
            string cPart = (absNum == 1 || absNum == 3) ? $"C{col}" : $"C[{col}]";
            reference = rPart + cPart;
        }

        if (!string.IsNullOrEmpty(sheet))
            reference = $"{sheet}!{reference}";
        return FormulaValue.Text(reference);
    }

    // ── SUBTOTAL ─────────────────────────────────────────────────────

    /// <summary>
    /// SUBTOTAL(function_num, ref1, ...). function_num 1-11 (and 101-111, which normally
    /// ignore manually hidden rows — treated identically here since we hold no hidden-row state).
    /// </summary>
    private static FormulaValue Subtotal(List<FormulaNode> a, FormulaContext c)
    {
        if (!FormulaHelper.TryEvalDouble(a[0], c, out double fnv, out var e)) return e;
        int f = (int)fnv;
        if (f >= 101 && f <= 111) f -= 100;
        if (f < 1 || f > 11) return FormulaValue.ErrorValue;

        var args = a.GetRange(1, a.Count - 1);

        if (f == 3) // COUNTA
        {
            var vals = new List<FormulaValue>();
            FormulaHelper.TryFlattenArgs(args, c, vals, out _);
            int cnt = 0; foreach (var v in vals) if (!v.IsBlank) cnt++;
            return FormulaValue.Number(cnt);
        }

        var nums = new List<double>();
        if (!FormulaHelper.TryCollectNumbers(args, c, nums, out var err)) return err;

        switch (f)
        {
            case 2: return FormulaValue.Number(nums.Count); // COUNT
            case 1: return nums.Count == 0 ? FormulaValue.ErrorDiv0 : FormulaValue.Number(nums.Sum() / nums.Count);
            case 4: return nums.Count == 0 ? FormulaValue.Zero : FormulaValue.Number(nums.Max());
            case 5: return nums.Count == 0 ? FormulaValue.Zero : FormulaValue.Number(nums.Min());
            case 6: { double p = 1; foreach (var n in nums) p *= n; return FormulaValue.Number(p); }
            case 9: { double s = 0; foreach (var n in nums) s += n; return FormulaValue.Number(s); }
            case 7: return Deviation(nums, sample: true, variance: false);   // STDEV
            case 8: return Deviation(nums, sample: false, variance: false);  // STDEVP
            case 10: return Deviation(nums, sample: true, variance: true);   // VAR
            case 11: return Deviation(nums, sample: false, variance: true);  // VARP
            default: return FormulaValue.ErrorValue;
        }
    }

    private static FormulaValue Deviation(List<double> nums, bool sample, bool variance)
    {
        int n = nums.Count;
        if (n == 0 || (sample && n == 1)) return FormulaValue.ErrorDiv0;
        double mean = nums.Sum() / n;
        double ss = 0; foreach (var v in nums) ss += (v - mean) * (v - mean);
        double var = ss / (sample ? n - 1 : n);
        return FormulaValue.Number(variance ? var : Math.Sqrt(var));
    }

    /// <summary>Column index (1-based) to letters, e.g. 3 → "C", 27 → "AA".</summary>
    private static string ColumnLetter(int col)
    {
        // Derive from the canonical single-cell reference (drops the trailing row digit).
        string s = CellReference.FromRowColumn(1, col);
        return s[..^1];
    }

    private static ArrayValue WrapSingle(FormulaValue v)
    {
        var arr = new ArrayValue(1, 1); arr[0] = v; return arr;
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
