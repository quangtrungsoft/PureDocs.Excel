using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents a worksheet in an Excel workbook.
/// </summary>
public sealed class Worksheet
{
    private readonly WorksheetPart _worksheetPart;
    private readonly Sheet _sheet;
    private readonly SharedStringManager _sharedStringManager;
    private readonly StyleManager _styleManager;
    private Workbook? _workbook;

    /// <summary>Sets the parent workbook reference (called by Workbook).</summary>
    internal void SetWorkbook(Workbook workbook) => _workbook = workbook;

    /// <summary>Gets the parent workbook, if available.</summary>
    internal Workbook? GetWorkbook() => _workbook;

    internal Worksheet(WorksheetPart worksheetPart, Sheet sheet, SharedStringManager sharedStringManager, StyleManager styleManager)
    {
        _worksheetPart = worksheetPart ?? throw new ArgumentNullException(nameof(worksheetPart));
        _sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
        _sharedStringManager = sharedStringManager ?? throw new ArgumentNullException(nameof(sharedStringManager));
        _styleManager = styleManager ?? throw new ArgumentNullException(nameof(styleManager));
    }

    /// <summary>
    /// Gets the style manager for this worksheet's workbook.
    /// </summary>
    internal StyleManager StyleManager => _styleManager;

    /// <summary>
    /// Exposes the underlying <see cref="WorksheetPart"/> for chart embedding.
    /// Called exclusively by <see cref="Charts.WorksheetChartExtensions"/>.
    /// </summary>
    internal WorksheetPart GetWorksheetPart() => _worksheetPart;

    private Formulas.CalcChain? _calcChain;

    /// <summary>Gets the calculation chain for dependency-aware recalc.</summary>
    public Formulas.CalcChain CalcChain => _calcChain ??= new Formulas.CalcChain();

    /// <summary>
    /// The workbook's named-range manager, or null when this worksheet is not
    /// attached to a workbook. Used to resolve named ranges during evaluation.
    /// </summary>
    internal Formulas.NamedRangeManager? NamedRanges => _workbook?.NamedRanges;

    /// <summary>
    /// This worksheet's 0-based index within the workbook (matches OpenXML localSheetId),
    /// or -1 when detached. Used to resolve sheet-scoped named ranges.
    /// </summary>
    internal int SheetIndex => _workbook?.Worksheets.IndexOf(this) ?? -1;

    /// <summary>
    /// Returns the spilled value for a cell that belongs to an active dynamic-array
    /// spill region (but holds no formula of its own), or null if the cell is not spilled.
    /// </summary>
    internal Formulas.FormulaValue? TryGetSpillValue(string reference)
    {
        if (_calcChain == null || string.IsNullOrEmpty(reference)) return null;
        var addr = Formulas.CellAddress.FromReference(reference);
        return _calcChain.Spills.GetSpillValue(addr);
    }

    /// <summary>
    /// Persists active dynamic-array spill regions into the OpenXML sheet: the anchor's
    /// formula is marked as an array formula spanning the spill range (<c>&lt;f t="array" ref="A1:A3"&gt;</c>),
    /// and cached values are written into every cell of the region so other tools read the
    /// spilled results. Returns the anchor cells that were written (for workbook-level
    /// dynamic-array metadata). Called during save.
    /// </summary>
    internal List<string> FlushSpillRanges()
    {
        var anchors = new List<string>();
        if (_calcChain == null) return anchors;

        foreach (var region in _calcChain.Spills.Regions)
        {
            var anchor = region.Anchor;
            int rows = region.Rows, cols = region.Columns;
            string anchorRef = anchor.ToReference();
            string endRef = CellReference.FromRowColumn(anchor.Row + rows - 1, anchor.Column + cols - 1);
            string spillRange = (rows == 1 && cols == 1) ? anchorRef : $"{anchorRef}:{endRef}";

            var anchorCell = GetCell(anchorRef).OpenXmlCell;
            if (anchorCell.CellFormula != null)
            {
                anchorCell.CellFormula.FormulaType = CellFormulaValues.Array;
                anchorCell.CellFormula.Reference = spillRange;
            }

            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols; c++)
                {
                    string cellRef = CellReference.FromRowColumn(anchor.Row + r, anchor.Column + c);
                    WriteCachedValue(GetCell(cellRef).OpenXmlCell, region.Data[r, c]);
                }

            anchors.Add(anchorRef);
        }
        return anchors;
    }

    /// <summary>Writes a computed FormulaValue as an OpenXML cell's cached value.</summary>
    private void WriteCachedValue(DocumentFormat.OpenXml.Spreadsheet.Cell oxCell, Formulas.FormulaValue val)
    {
        if (val.IsNumber)
        {
            oxCell.CellValue = new CellValue(val.NumberValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
            oxCell.DataType = new EnumValue<CellValues>(CellValues.Number);
        }
        else if (val.IsBoolean)
        {
            oxCell.CellValue = new CellValue(val.BooleanValue ? "1" : "0");
            oxCell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
        }
        else if (val.IsError)
        {
            oxCell.CellValue = new CellValue(val.ToString());
            oxCell.DataType = new EnumValue<CellValues>(CellValues.Error);
        }
        else if (val.IsText)
        {
            int idx = _sharedStringManager.AddOrGetString(val.TextValue);
            oxCell.CellValue = new CellValue(idx.ToString());
            oxCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
        }
        else
        {
            oxCell.CellValue = null;
            oxCell.DataType = null;
        }
    }

    private bool _suspendCalcTracking;

    /// <summary>
    /// Routes a cell content change into the calc chain so that a later
    /// <see cref="SmartRecalculate"/> re-evaluates the affected cells.
    /// Only acts once a calc chain exists (i.e. after the first recalc), so
    /// worksheets that never recalc pay no tracking cost. Called by <see cref="Cell"/>.
    /// </summary>
    internal void NotifyCellContentChanged(string reference, string? formula)
    {
        if (_calcChain == null || _suspendCalcTracking || string.IsNullOrEmpty(reference))
            return;

        var address = Formulas.CellAddress.FromReference(reference);
        if (formula != null)
            _calcChain.SetFormula(address, formula);   // parse, update deps, mark dirty
        else
            _calcChain.RemoveFormula(address);          // drop any stale formula tracking + mark dependents dirty
    }

    /// <summary>
    /// Suspends per-cell recalc tracking. Use around bulk edits (e.g. large
    /// <see cref="Range.SetValues"/>) to avoid marking dependents dirty on every write;
    /// call <see cref="ResumeRecalcTracking"/> afterwards. The next
    /// <see cref="SmartRecalculate"/> then performs a full rebuild.
    /// </summary>
    public void SuspendRecalcTracking() => _suspendCalcTracking = true;

    /// <summary>
    /// Resumes per-cell recalc tracking and schedules a full recalc on the next
    /// <see cref="SmartRecalculate"/>, since edits made while suspended were not tracked.
    /// </summary>
    public void ResumeRecalcTracking()
    {
        _suspendCalcTracking = false;
        _calcChain?.MarkAllDirty();
    }

    /// <summary>Gets the underlying SheetData element for formula scanning.</summary>
    internal DocumentFormat.OpenXml.Spreadsheet.SheetData? GetSheetData()
        => _worksheetPart.Worksheet.GetFirstChild<SheetData>();

    /// <summary>
    /// Recalculates only dirty cells using the dependency graph.
    /// Much faster than CalculateFormulas() for incremental changes.
    /// First call does a full build; subsequent calls are incremental.
    /// </summary>
    public int SmartRecalculate()
    {
        return CalcChain.Recalculate(this);
    }

    /// <summary>
    /// Evaluates a formula string against this worksheet.
    /// </summary>
    /// <param name="formula">The formula to evaluate (with or without leading '=').</param>
    /// <returns>The computed value (object? for backward compatibility).</returns>
    public object? EvaluateFormula(string formula)
    {
        return Formulas.FormulaEvaluator.Evaluate(formula, this, NamedRanges).ToObject();
    }

    /// <summary>
    /// Evaluates a formula and returns the typed FormulaValue.
    /// </summary>
    internal Formulas.FormulaValue EvaluateFormulaValue(string formula)
    {
        return Formulas.FormulaEvaluator.Evaluate(formula, this, NamedRanges);
    }

    /// <summary>
    /// Recalculates all formulas in this worksheet by evaluating them and
    /// storing the cached values.
    /// </summary>
    public void CalculateFormulas()
    {
        var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;

        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var oxCell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                if (oxCell.CellFormula != null && !string.IsNullOrEmpty(oxCell.CellFormula.Text))
                {
                    var result = Formulas.FormulaEvaluator.Evaluate(oxCell.CellFormula.Text, this, NamedRanges);

                    if (result.IsError)
                        continue; // Leave error formulas as-is

                    if (result.IsNumber)
                    {
                        oxCell.CellValue = new CellValue(result.NumberValue.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        oxCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Number);
                    }
                    else if (result.IsBoolean)
                    {
                        oxCell.CellValue = new CellValue(result.BooleanValue ? "1" : "0");
                        oxCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.Boolean);
                    }
                    else if (result.IsText)
                    {
                        int idx = _sharedStringManager.AddOrGetString(result.TextValue);
                        oxCell.CellValue = new CellValue(idx.ToString());
                        oxCell.DataType = new DocumentFormat.OpenXml.EnumValue<CellValues>(CellValues.SharedString);
                    }
                }
            }
        }
    }

    /// <summary>
    /// Gets or sets the name of the worksheet.
    /// </summary>
    public string Name
    {
        get => _sheet.Name?.Value ?? string.Empty;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(value));
            _sheet.Name = value;
        }
    }

    /// <summary>
    /// Gets a cell by Excel reference (e.g., "A1").
    /// Creates the cell if it doesn't exist (use for writing).
    /// </summary>
    /// <remarks>
    /// This method creates row/cell elements if they don't exist.
    /// For read-only access without modifying the document, use <see cref="TryGetCell"/> instead.
    /// </remarks>
    public Cell GetCell(string cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
            throw new ArgumentException("Cell reference cannot be null or empty.", nameof(cellReference));

        var worksheetPart = _worksheetPart;
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>() 
            ?? worksheetPart.Worksheet.AppendChild(new SheetData());

        // Canonicalize (strip '$', normalize case) so "$C$2"/"c2"/"C2" all map to
        // the same stored cell — matching Excel addressing and the ordinal-sorted layout.
        CellReference.Parse(cellReference, out int rowIndex, out int colIndex);
        string canonical = CellReference.FromRowColumn(rowIndex, colIndex);

        var row = FindOrCreateRow(sheetData, rowIndex);
        var cell = FindOrCreateCell(row, canonical);
        return new Cell(cell, _sharedStringManager, _styleManager, this);
    }

    /// <summary>
    /// Tries to get an existing cell by Excel reference (e.g., "A1").
    /// Returns null if the cell doesn't exist (read-only, doesn't modify document).
    /// </summary>
    /// <remarks>
    /// Use this method for read-only access to avoid creating empty rows/cells in the document.
    /// For write access, use <see cref="GetCell(string)"/> instead.
    /// </remarks>
    public Cell? TryGetCell(string cellReference)
    {
        if (string.IsNullOrWhiteSpace(cellReference))
            throw new ArgumentException("Cell reference cannot be null or empty.", nameof(cellReference));

        var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return null;

        CellReference.Parse(cellReference, out int rowIndex, out int colIndex);
        string canonical = CellReference.FromRowColumn(rowIndex, colIndex);

        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row == null) return null;

        var cell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == canonical);

        if (cell == null) return null;

        return new Cell(cell, _sharedStringManager, _styleManager, this);
    }

    /// <summary>
    /// Tries to get an existing cell by row and column indices (1-based).
    /// Returns null if the cell doesn't exist (read-only, doesn't modify document).
    /// </summary>
    public Cell? TryGetCell(int row, int column)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));

        string cellReference = CellReference.FromRowColumn(row, column);
        return TryGetCell(cellReference);
    }

    /// <summary>
    /// Gets the value of a cell by reference (e.g. "A1").
    /// Used by formula engine. Returns null if empty.
    /// </summary>
    public object? GetCellValue(string cellReference)
    {
        // Optimized path: don't create elements if they don't exist
        var worksheetPart = _worksheetPart;
        var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return null;

        CellReference.Parse(cellReference, out int rowIndex, out int colIndex);
        string canonical = CellReference.FromRowColumn(rowIndex, colIndex);

        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == rowIndex);
        if (row == null) return null;

        var cell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == canonical);

        if (cell == null) return null;

        return new Cell(cell, _sharedStringManager, _styleManager).GetValue();
    }

    /// <summary>
    /// Gets values for a range of cells.
    /// </summary>
    public Formulas.FormulaValue GetRangeValues(Formulas.CellAddress start, Formulas.CellAddress end)
    {
         // This is a helper for 3D refs or external calls
         int rows = end.Row - start.Row + 1;
         int cols = end.Column - start.Column + 1;
         var arr = new Formulas.ArrayValue(rows, cols);
         
         for(int r=0; r<rows; r++)
         {
             for(int c=0; c<cols; c++)
             {
                 string refStr = CellReference.FromRowColumn(start.Row + r, start.Column + c);
                 var val = GetCellValue(refStr);
                 arr[r, c] = Formulas.FormulaValue.FromObject(val);
             }
         }
         return Formulas.FormulaValue.Array(arr);
    }

    /// <summary>
    /// Gets a cell by row and column indices (1-based).
    /// </summary>
    public Cell GetCell(int row, int column)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));

        string cellReference = CellReference.FromRowColumn(row, column);
        return GetCell(cellReference);
    }

    /// <summary>
    /// Cell indexer - sheet["A1"] or sheet[1, 1]
    /// </summary>
    public Cell this[string cellReference] => GetCell(cellReference);
    public Cell this[int row, int column] => GetCell(row, column);

    /// <summary>
    /// Gets a range of cells.
    /// </summary>
    public Range GetRange(string rangeReference)
    {
        if (string.IsNullOrWhiteSpace(rangeReference))
            throw new ArgumentException("Range reference cannot be null or empty.", nameof(rangeReference));
        return new Range(this, rangeReference);
    }

    /// <summary>
    /// Sets the width of a column.
    /// </summary>
    public void SetColumnWidth(int columnIndex, double width)
    {
        if (columnIndex < 1) throw new ArgumentOutOfRangeException(nameof(columnIndex));
        if (width <= 0) throw new ArgumentOutOfRangeException(nameof(width));

        var worksheet = _worksheetPart.Worksheet;
        var columns = worksheet.GetFirstChild<Columns>();

        if (columns == null)
        {
            columns = new Columns();
            worksheet.InsertBefore(columns, worksheet.GetFirstChild<SheetData>());
        }

        var existingColumn = columns.Elements<Column>()
            .FirstOrDefault(c => c.Min == columnIndex && c.Max == columnIndex);
        existingColumn?.Remove();

        var column = new Column
        {
            Min = (uint)columnIndex,
            Max = (uint)columnIndex,
            Width = width,
            CustomWidth = true
        };
        columns.Append(column);
    }

    /// <summary>
    /// Auto-fits a column based on content.
    /// </summary>
    public void AutoFitColumn(int columnIndex)
    {
        if (columnIndex < 1) throw new ArgumentOutOfRangeException(nameof(columnIndex));

        double maxWidth = 10;

        var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData != null)
        {
            foreach (var row in sheetData.Elements<Row>())
            {
                var cell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                    .FirstOrDefault(c =>
                    {
                        // Null check: CellReference can be null in valid OpenXml documents
                        if (c.CellReference?.Value == null) return false;
                        CellReference.Parse(c.CellReference.Value, out _, out int col);
                        return col == columnIndex;
                    });

                if (cell != null)
                {
                    var cellWrapper = new Cell(cell, _sharedStringManager, _styleManager);
                    string text = cellWrapper.GetText();
                    double textWidth = text.Length * 1.2 + 2;
                    maxWidth = Math.Max(maxWidth, textWidth);
                }
            }
        }

        SetColumnWidth(columnIndex, Math.Min(maxWidth, 255));
    }

    /// <summary>
    /// Sets the height of a row.
    /// </summary>
    public void SetRowHeight(int rowIndex, double height)
    {
        if (rowIndex < 1) throw new ArgumentOutOfRangeException(nameof(rowIndex));
        if (height <= 0) throw new ArgumentOutOfRangeException(nameof(height));

        var sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>()
            ?? _worksheetPart.Worksheet.AppendChild(new SheetData());

        var row = FindOrCreateRow(sheetData, rowIndex);
        row.Height = height;
        row.CustomHeight = true;
    }

    /// <summary>
    /// Merges a range of cells.
    /// </summary>
    public void MergeCells(string rangeReference)
    {
        if (string.IsNullOrWhiteSpace(rangeReference))
            throw new ArgumentException("Range reference cannot be null or empty.", nameof(rangeReference));

        var worksheet = _worksheetPart.Worksheet;
        var mergeCells = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.MergeCells>();

        if (mergeCells == null)
        {
            mergeCells = new DocumentFormat.OpenXml.Spreadsheet.MergeCells();
            // MergeCells must come after SheetData
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null)
                worksheet.InsertAfter(mergeCells, sheetData);
            else
                worksheet.Append(mergeCells);
        }

        mergeCells.Append(new MergeCell { Reference = rangeReference });
    }

    /// <summary>
    /// Unmerges cells in the specified range.
    /// </summary>
    public void UnmergeCells(string rangeReference)
    {
        if (string.IsNullOrWhiteSpace(rangeReference))
            throw new ArgumentException("Range reference cannot be null or empty.", nameof(rangeReference));

        var mergeCells = _worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.MergeCells>();
        if (mergeCells == null) return;

        var mergeCell = mergeCells.Elements<MergeCell>()
            .FirstOrDefault(mc => mc.Reference?.Value == rangeReference);

        mergeCell?.Remove();

        if (!mergeCells.HasChildren)
            mergeCells.Remove();
    }

    /// <summary>
    /// Freezes panes at the specified row and column.
    /// </summary>
    /// <param name="row">Number of rows to freeze (0 = no row freeze).</param>
    /// <param name="column">Number of columns to freeze (0 = no column freeze).</param>
    /// <remarks>
    /// FreezePanes(3, 0) freezes the first 3 rows (TopLeftCell = A4).
    /// FreezePanes(0, 2) freezes the first 2 columns (TopLeftCell = C1).
    /// FreezePanes(3, 2) freezes 3 rows and 2 columns (TopLeftCell = C4).
    /// </remarks>
    public void FreezePanes(int row, int column)
    {
        if (row < 0) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 0) throw new ArgumentOutOfRangeException(nameof(column));

        var worksheet = _worksheetPart.Worksheet;
        var sheetViews = worksheet.GetFirstChild<SheetViews>();
        if (sheetViews == null)
        {
            sheetViews = new SheetViews();
            worksheet.InsertBefore(sheetViews, worksheet.GetFirstChild<SheetFormatProperties>()
                ?? worksheet.GetFirstChild<Columns>()
                ?? (OpenXmlElement?)worksheet.GetFirstChild<SheetData>());
        }

        var sheetView = sheetViews.GetFirstChild<SheetView>()
            ?? sheetViews.AppendChild(new SheetView { WorkbookViewId = 0 });

        // Remove existing pane
        sheetView.RemoveAllChildren<Pane>();

        if (row == 0 && column == 0) return; // No freeze

        // TopLeftCell is the first unfrozen cell
        // If row=3, column=0: TopLeftCell = A4 (row 4, col 1)
        // If row=0, column=2: TopLeftCell = C1 (row 1, col 3)
        // If row=3, column=2: TopLeftCell = C4 (row 4, col 3)
        int topLeftRow = row > 0 ? row + 1 : 1;
        int topLeftCol = column > 0 ? column + 1 : 1;

        var pane = new Pane
        {
            State = PaneStateValues.Frozen,
            TopLeftCell = CellReference.FromRowColumn(topLeftRow, topLeftCol)
        };

        if (row > 0) pane.VerticalSplit = row;
        if (column > 0) pane.HorizontalSplit = column;

        string activePane;
        if (row > 0 && column > 0)
            activePane = "bottomRight";
        else if (row > 0)
            activePane = "bottomLeft";
        else
            activePane = "topRight";

        pane.ActivePane = activePane switch
        {
            "bottomRight" => PaneValues.BottomRight,
            "bottomLeft" => PaneValues.BottomLeft,
            "topRight" => PaneValues.TopRight,
            _ => PaneValues.BottomLeft
        };

        sheetView.Append(pane);
    }

    /// <summary>
    /// Adds a dropdown data validation to the specified cell range using a list formula.
    /// </summary>
    /// <param name="cellRange">The cell range to apply validation (e.g., "F2:F1000").</param>
    /// <param name="listFormula">The formula for the dropdown list (e.g., "'Danh mục'!$A$2:$A$10").</param>
    public void AddDropdownValidation(string cellRange, string listFormula)
    {
        if (string.IsNullOrWhiteSpace(cellRange))
            throw new ArgumentException("Cell range cannot be null or empty.", nameof(cellRange));
        if (string.IsNullOrWhiteSpace(listFormula))
            throw new ArgumentException("List formula cannot be null or empty.", nameof(listFormula));

        var worksheet = _worksheetPart.Worksheet;
        var dataValidations = worksheet.GetFirstChild<DataValidations>();

        if (dataValidations == null)
        {
            dataValidations = new DataValidations();

            // DataValidations should come after AutoFilter, MergeCells, or SheetData
            var autoFilter = worksheet.GetFirstChild<AutoFilter>();
            var mergeCells = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.MergeCells>();
            var sheetData = worksheet.GetFirstChild<SheetData>();

            if (autoFilter != null)
                worksheet.InsertAfter(dataValidations, autoFilter);
            else if (mergeCells != null)
                worksheet.InsertAfter(dataValidations, mergeCells);
            else if (sheetData != null)
                worksheet.InsertAfter(dataValidations, sheetData);
            else
                worksheet.Append(dataValidations);
        }

        var validation = new DataValidation
        {
            Type = DataValidationValues.List,
            AllowBlank = true,
            ShowInputMessage = true,
            ShowErrorMessage = true,
            SequenceOfReferences = new ListValue<StringValue> { InnerText = cellRange },
            Formula1 = new Formula1(listFormula)
        };

        dataValidations.Append(validation);
        dataValidations.Count = (uint)dataValidations.ChildElements.Count;
    }

    /// <summary>
    /// Sets the visibility state of this worksheet.
    /// </summary>
    /// <param name="state">The sheet state (Visible, Hidden, VeryHidden).</param>
    public void SetSheetVisibility(SheetStateValues state)
    {
        _sheet.State = state;
    }

    /// <summary>
    /// Sets auto-filter on the specified range.
    /// </summary>
    public void SetAutoFilter(string rangeReference)
    {
        if (string.IsNullOrWhiteSpace(rangeReference))
            throw new ArgumentException("Range reference cannot be null or empty.", nameof(rangeReference));

        var worksheet = _worksheetPart.Worksheet;

        // Remove existing auto filter
        var existingFilter = worksheet.GetFirstChild<AutoFilter>();
        existingFilter?.Remove();

        var autoFilter = new AutoFilter { Reference = rangeReference };

        // AutoFilter must come after MergeCells or SheetData
        var mergeCells = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.MergeCells>();
        if (mergeCells != null)
            worksheet.InsertAfter(autoFilter, mergeCells);
        else
        {
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null)
                worksheet.InsertAfter(autoFilter, sheetData);
            else
                worksheet.Append(autoFilter);
        }
    }

    private static Row FindOrCreateRow(SheetData sheetData, int rowIndex)
    {
        var row = sheetData.Elements<Row>()
            .FirstOrDefault(r => r.RowIndex?.Value == rowIndex);

        if (row == null)
        {
            row = new Row { RowIndex = (uint)rowIndex };
            Row? refRow = null;
            foreach (var existingRow in sheetData.Elements<Row>())
            {
                if (existingRow.RowIndex!.Value > rowIndex)
                {
                    refRow = existingRow;
                    break;
                }
            }
            if (refRow != null)
                sheetData.InsertBefore(row, refRow);
            else
                sheetData.Append(row);
        }
        return row;
    }

    private static DocumentFormat.OpenXml.Spreadsheet.Cell FindOrCreateCell(Row row, string cellReference)
    {
        var cell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()
            .FirstOrDefault(c => c.CellReference?.Value == cellReference);

        if (cell == null)
        {
            cell = new DocumentFormat.OpenXml.Spreadsheet.Cell { CellReference = cellReference };
            DocumentFormat.OpenXml.Spreadsheet.Cell? refCell = null;
            foreach (var existingCell in row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                if (string.Compare(existingCell.CellReference?.Value, cellReference, StringComparison.Ordinal) > 0)
                {
                    refCell = existingCell;
                    break;
                }
            }
            if (refCell != null)
                row.InsertBefore(cell, refCell);
            else
                row.Append(cell);
        }
        return cell;
    }
}