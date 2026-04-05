using System.Collections.Concurrent;

namespace TVE.PureDocs.Excel.Formulas;

/// <summary>
/// Evaluation context: resolves cell/range values, dispatches functions.
/// Supports cross-sheet references and named ranges.
/// Supports both single-threaded (HashSet) and thread-safe (ConcurrentDictionary) 
/// circular reference detection.
/// </summary>
internal sealed class FormulaContext
{
    private readonly Worksheet _worksheet;
    private readonly FunctionRegistry _functions;
    private readonly HashSet<string>? _evaluatingCells;
    private readonly ConcurrentDictionary<string, byte>? _evaluatingCellsThreadSafe;
    private readonly NamedRangeManager? _namedRanges;

    public Worksheet Worksheet => _worksheet;
    
    /// <summary>Gets evaluating cells set (for single-threaded use).</summary>
    public HashSet<string>? EvaluatingCells => _evaluatingCells;

    /// <summary>
    /// Gets or sets the row of the formula cell being evaluated.
    /// Used for implicit intersection (@) operator.
    /// </summary>
    public int FormulaRow { get; set; }

    /// <summary>
    /// Gets or sets the column of the formula cell being evaluated.
    /// Used for implicit intersection (@) operator.
    /// </summary>
    public int FormulaCol { get; set; }

    /// <summary>
    /// Creates a FormulaContext for single-threaded evaluation.
    /// </summary>
    public FormulaContext(Worksheet worksheet, HashSet<string>? evaluatingCells = null,
        NamedRangeManager? namedRanges = null)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _functions = FunctionRegistry.Default;
        _evaluatingCells = evaluatingCells ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        _evaluatingCellsThreadSafe = null;
        _namedRanges = namedRanges;
        FormulaRow = 0;
        FormulaCol = 0;
    }

    /// <summary>
    /// Creates a FormulaContext for thread-safe parallel evaluation.
    /// Uses ConcurrentDictionary for circular reference detection across threads.
    /// </summary>
    public FormulaContext(Worksheet worksheet, ConcurrentDictionary<string, byte> threadSafeEvaluatingCells,
        NamedRangeManager? namedRanges = null)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _functions = FunctionRegistry.Default;
        _evaluatingCells = null;
        _evaluatingCellsThreadSafe = threadSafeEvaluatingCells ?? throw new ArgumentNullException(nameof(threadSafeEvaluatingCells));
        _namedRanges = namedRanges;
        FormulaRow = 0;
        FormulaCol = 0;
    }

    /// <summary>
    /// Creates a FormulaContext with formula position for implicit intersection.
    /// </summary>
    public FormulaContext(Worksheet worksheet, int formulaRow, int formulaCol, 
        HashSet<string>? evaluatingCells = null, NamedRangeManager? namedRanges = null)
        : this(worksheet, evaluatingCells, namedRanges)
    {
        FormulaRow = formulaRow;
        FormulaCol = formulaCol;
    }

    /// <summary>Gets the value of a single cell. Recursively evaluates formulas.</summary>
    public FormulaValue GetCellValue(string cellReference)
    {
        string norm = cellReference.Replace("$", "").ToUpperInvariant();

        // Circular reference detection (thread-safe or single-threaded)
        bool added;
        if (_evaluatingCellsThreadSafe != null)
        {
            added = _evaluatingCellsThreadSafe.TryAdd(norm, 0);
        }
        else
        {
            added = _evaluatingCells!.Add(norm);
        }

        if (!added)
            return FormulaValue.ErrorRef;

        try
        {
            var cell = _worksheet.GetCell(cellReference);

            if (cell.HasFormula)
            {
                string formula = cell.GetFormula()!;
                // For thread-safe mode, pass the concurrent dictionary
                if (_evaluatingCellsThreadSafe != null)
                {
                    return EvaluateInternalThreadSafe(formula, _worksheet, _evaluatingCellsThreadSafe, _namedRanges);
                }
                return FormulaEvaluator.EvaluateInternal(formula, _worksheet, _evaluatingCells!, _namedRanges);
            }

            return FormulaValue.FromObject(cell.GetValue());
        }
        catch (FormulaException)
        {
            return FormulaValue.ErrorValue;
        }
        finally
        {
            if (_evaluatingCellsThreadSafe != null)
            {
                _evaluatingCellsThreadSafe.TryRemove(norm, out _);
            }
            else
            {
                _evaluatingCells!.Remove(norm);
            }
        }
    }

    /// <summary>Internal thread-safe evaluate for parallel recalculation.</summary>
    private static FormulaValue EvaluateInternalThreadSafe(string formula, Worksheet worksheet,
        ConcurrentDictionary<string, byte> evaluatingCells, NamedRangeManager? namedRanges)
    {
        if (string.IsNullOrWhiteSpace(formula))
            return FormulaValue.Blank;

        formula = formula.TrimStart('=');

        try
        {
            var ast = FormulaEvaluator.GetOrParseAst(formula);
            var context = new FormulaContext(worksheet, evaluatingCells, namedRanges);
            return ast.Evaluate(context);
        }
        catch (FormulaException ex)
        {
            return FormulaValue.Error(FormulaValue.ErrorFromString(ex.Message));
        }
        catch (Exception)
        {
            return FormulaValue.ErrorValue;
        }
    }

    /// <summary>Gets range values as ArrayValue.</summary>
    public FormulaValue GetRangeValues(string startRef, string endRef)
    {
        CellReference.Parse(startRef, out int sr, out int sc);
        CellReference.Parse(endRef, out int er, out int ec);

        int rows = er - sr + 1, cols = ec - sc + 1;
        var arr = new ArrayValue(rows, cols);

        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
            {
                string cellRef = CellReference.FromRowColumn(sr + r, sc + c);
                arr[r, c] = GetCellValue(cellRef);
            }

        return FormulaValue.Array(arr);
    }

    /// <summary>Gets range values as 2D array (for matrix operations).</summary>
    public FormulaValue[,] GetRangeValues2D(string startRef, string endRef)
    {
        CellReference.Parse(startRef, out int sr, out int sc);
        CellReference.Parse(endRef, out int er, out int ec);

        int rows = er - sr + 1, cols = ec - sc + 1;
        var result = new FormulaValue[rows, cols];

        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                result[r, c] = GetCellValue(CellReference.FromRowColumn(sr + r, sc + c));

        return result;
    }

    /// <summary>Gets range bounds.</summary>
    public (int startRow, int startCol, int endRow, int endCol) GetRangeBounds(string startRef, string endRef)
    {
        CellReference.Parse(startRef, out int sr, out int sc);
        CellReference.Parse(endRef, out int er, out int ec);
        return (sr, sc, er, ec);
    }

    /// <summary>Gets a cell value from a different sheet.</summary>
    public FormulaValue GetSheetCellValue(string sheetName, string cellReference)
    {
        try
        {
            var workbook = _worksheet.GetWorkbook();
            if (workbook == null) return FormulaValue.ErrorRef;

            var sheet = workbook.Worksheets[sheetName];
            if (sheet == null) return FormulaValue.ErrorRef;

            // Create context for the target sheet, sharing evaluatingCells for circular ref detection
            // Use appropriate constructor based on thread-safe mode
            FormulaContext ctx;
            if (_evaluatingCellsThreadSafe != null)
            {
                ctx = new FormulaContext(sheet, _evaluatingCellsThreadSafe, _namedRanges);
            }
            else
            {
                ctx = new FormulaContext(sheet, _evaluatingCells, _namedRanges);
            }
            return ctx.GetCellValue(cellReference);
        }
        catch
        {
            return FormulaValue.ErrorRef;
        }
    }

    /// <summary>Gets range values from a different sheet.</summary>
    public FormulaValue GetSheetRangeValues(string sheetName, string startRef, string endRef)
    {
        try
        {
            var workbook = _worksheet.GetWorkbook();
            if (workbook == null) return FormulaValue.ErrorRef;

            var sheet = workbook.Worksheets[sheetName];
            if (sheet == null) return FormulaValue.ErrorRef;

            // Use appropriate constructor based on thread-safe mode
            FormulaContext ctx;
            if (_evaluatingCellsThreadSafe != null)
            {
                ctx = new FormulaContext(sheet, _evaluatingCellsThreadSafe, _namedRanges);
            }
            else
            {
                ctx = new FormulaContext(sheet, _evaluatingCells, _namedRanges);
            }
            return ctx.GetRangeValues(startRef, endRef);
        }
        catch
        {
            return FormulaValue.ErrorRef;
        }
    }

    /// <summary>Resolves a named range reference.</summary>
    public FormulaValue ResolveNamedRange(string name)
    {
        if (_namedRanges == null || !_namedRanges.TryGet(name, out var range))
            return FormulaValue.ErrorName;

        var (sheetName, startRef, endRef) = range.ParseReference();

        // Determine which worksheet to read from
        Worksheet targetSheet = _worksheet;
        if (sheetName != null)
        {
            try
            {
                var workbook = _worksheet.GetWorkbook();
                if (workbook != null)
                {
                    var sheet = workbook.Worksheets[sheetName];
                    if (sheet != null) targetSheet = sheet;
                    else return FormulaValue.ErrorRef;
                }
            }
            catch { return FormulaValue.ErrorRef; }
        }

        var ctx = new FormulaContext(targetSheet, _evaluatingCells, _namedRanges);

        if (endRef != null)
            return ctx.GetRangeValues(startRef, endRef);
        else
            return ctx.GetCellValue(startRef);
    }

    /// <summary>Evaluates a function by name.</summary>
    public FormulaValue EvaluateFunction(string functionName, List<FormulaNode> arguments)
    {
        return _functions.Execute(functionName, arguments, this);
    }


    /// <summary>Table manager for structured references.</summary>
    public TableManager? TableManager { get; set; }

    /// <summary>Resolves a structured reference (e.g., Table1[Column1]).</summary>
    public FormulaValue ResolveStructuredReference(string tableName, string specifier)
    {
        if (TableManager == null) return FormulaValue.ErrorRef;

        var result = TableManager.Resolve(tableName, specifier, FormulaRow);
        if (result == null) return FormulaValue.ErrorRef;

        var val = result.Value;
        
        // If it refers to the same sheet/table, we can just resolve range normally.
        // Assuming tables live on the current worksheet for simplicity or we need sheet awareness in TableDefinition.
        // For now, let's assume simple single-sheet tables or global table names.
        
        // Convert row/col to strictly A1-style reference (e.g. A1:B10) won't work easily if we don't know the sheet.
        // BUT, since we are inside FormulaContext, we can try to map coordinates to cells if the table is on this sheet.
        
        // Optimization: Create an ArrayValue directly from the coordinates?
        // Or resolve via GetCellValue loop.
        
        int rows = val.EndRow - val.StartRow + 1;
        int cols = val.EndCol - val.StartCol + 1;
        
        // If single cell
        if (rows == 1 && cols == 1)
        {
            return GetCellValue(CellReference.FromRowColumn(val.StartRow, val.StartCol));
        }

        var arr = new ArrayValue(rows, cols);
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                arr[r, c] = GetCellValue(CellReference.FromRowColumn(val.StartRow + r, val.StartCol + c));
            }
        }
        return FormulaValue.Array(arr);
    }

    /// <summary>Resolves a 3D cell reference (Sheet1:Sheet3!A1).</summary>
    public FormulaValue Resolve3DCell(string startSheet, string endSheet, string cellRef)
    {
        var wb = _worksheet.GetWorkbook();
        if (wb == null) return FormulaValue.ErrorRef;
        return ThreeDReference.Resolve3DCell(startSheet, endSheet, cellRef, wb);
    }

    /// <summary>Resolves a 3D range reference (Sheet1:Sheet3!A1:B5).</summary>
    public FormulaValue Resolve3DRange(string startSheet, string endSheet, string startRef, string endRef)
    {
        var wb = _worksheet.GetWorkbook();
        if (wb == null) return FormulaValue.ErrorRef;
        return ThreeDReference.Resolve3DRange(startSheet, endSheet, startRef, endRef, wb);
    }
}
