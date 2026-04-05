namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Handles 3D references: Sheet1:Sheet3!A1 → aggregates A1 across Sheet1, Sheet2, Sheet3.
/// Used with aggregate functions like SUM(Sheet1:Sheet3!A1).
/// </summary>
public static class ThreeDReference
{
    /// <summary>
    /// Resolves a 3D reference to a list of values from multiple sheets.
    /// For Sheet1:Sheet3!A1 → gets A1 from Sheet1, Sheet2, Sheet3.
    /// For Sheet1:Sheet3!A1:B5 → gets the range A1:B5 from each sheet.
    /// </summary>
    public static FormulaValue Resolve3DCell(
        string startSheet, string endSheet, string cellRef,
        Excel.Workbook workbook)
    {
        var sheets = GetSheetRange(startSheet, endSheet, workbook);
        if (sheets == null)
            return FormulaValue.ErrorRef;

        var values = new List<FormulaValue>(sheets.Count);
        foreach (var sheet in sheets)
        {
            var value = GetCellValueFromSheet(sheet, cellRef);
            values.Add(value);
        }

        return FormulaValue.Array(ArrayValue.FromList(values));
    }

    /// <summary>
    /// Resolves a 3D range reference to a combined array of values.
    /// For Sheet1:Sheet3!A1:B2 → stacks all matching ranges vertically.
    /// </summary>
    public static FormulaValue Resolve3DRange(
        string startSheet, string endSheet,
        string startRef, string endRef,
        Excel.Workbook workbook)
    {
        var sheets = GetSheetRange(startSheet, endSheet, workbook);
        if (sheets == null)
            return FormulaValue.ErrorRef;

        CellReference.Parse(startRef, out int sr, out int sc);
        CellReference.Parse(endRef, out int er, out int ec);

        int rangeRows = er - sr + 1;
        int rangeCols = ec - sc + 1;
        int totalRows = rangeRows * sheets.Count;

        var result = new ArrayValue(totalRows, rangeCols);
        int rowOffset = 0;

        foreach (var sheet in sheets)
        {
            for (int r = sr; r <= er; r++)
            {
                for (int c = sc; c <= ec; c++)
                {
                    string cellRef = CellReference.FromRowColumn(r, c);
                    var value = GetCellValueFromSheet(sheet, cellRef);
                    result[rowOffset + (r - sr), c - sc] = value;
                }
            }
            rowOffset += rangeRows;
        }

        return FormulaValue.Array(result);
    }

    /// <summary>
    /// Returns the contiguous sequence of sheets from startSheet to endSheet.
    /// Returns null if either sheet is not found.
    /// </summary>
    /// <remarks>
    /// Sheet names may be quoted (e.g., 'My Sheet') from the lexer.
    /// This method strips surrounding quotes for proper comparison.
    /// </remarks>
    private static List<Excel.Worksheet>? GetSheetRange(
        string startSheet, string endSheet, Excel.Workbook workbook)
    {
        // Strip surrounding quotes if present (lexer may include them)
        startSheet = StripQuotes(startSheet);
        endSheet = StripQuotes(endSheet);
        
        var worksheets = workbook.Worksheets;
        int startIdx = -1, endIdx = -1;

        for (int i = 0; i < worksheets.Count; i++)
        {
            var name = worksheets[i].Name;
            if (string.Equals(name, startSheet, StringComparison.OrdinalIgnoreCase))
                startIdx = i;
            if (string.Equals(name, endSheet, StringComparison.OrdinalIgnoreCase))
                endIdx = i;
        }

        if (startIdx < 0 || endIdx < 0)
            return null;

        // Ensure correct order
        if (startIdx > endIdx)
            (startIdx, endIdx) = (endIdx, startIdx);

        var result = new List<Excel.Worksheet>(endIdx - startIdx + 1);
        for (int i = startIdx; i <= endIdx; i++)
            result.Add(worksheets[i]);

        return result;
    }

    /// <summary>
    /// Strips surrounding single quotes from a sheet name.
    /// Also handles escaped quotes ('') inside the name.
    /// </summary>
    private static string StripQuotes(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName)) return sheetName;
        
        // Strip surrounding single quotes
        if (sheetName.StartsWith('\'') && sheetName.EndsWith('\'') && sheetName.Length >= 2)
        {
            sheetName = sheetName[1..^1];
            // Unescape doubled quotes ('') → (')
            sheetName = sheetName.Replace("''", "'");
        }
        
        return sheetName;
    }

    private static FormulaValue GetCellValueFromSheet(Excel.Worksheet sheet, string cellRef)
    {
        try
        {
            var cellValue = sheet.GetCellValue(cellRef);
            return FormulaValue.FromObject(cellValue);
        }
        catch
        {
            return FormulaValue.Blank;
        }
    }
}
