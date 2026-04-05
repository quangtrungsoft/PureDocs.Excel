using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Represents a cell in a worksheet.
/// </summary>
public sealed class Cell
{
    private readonly DocumentFormat.OpenXml.Spreadsheet.Cell _cell;
    private readonly SharedStringManager _sharedStringManager;
    private readonly StyleManager _styleManager;

    internal Cell(DocumentFormat.OpenXml.Spreadsheet.Cell cell, SharedStringManager sharedStringManager, StyleManager styleManager)
    {
        _cell = cell ?? throw new ArgumentNullException(nameof(cell));
        _sharedStringManager = sharedStringManager ?? throw new ArgumentNullException(nameof(sharedStringManager));
        _styleManager = styleManager ?? throw new ArgumentNullException(nameof(styleManager));
    }

    /// <summary>
    /// Gets the cell reference (e.g., "A1").
    /// </summary>
    public string Reference => _cell.CellReference?.Value ?? string.Empty;

    /// <summary>
    /// Gets the row index (1-based).
    /// </summary>
    public int Row
    {
        get
        {
            CellReference.Parse(Reference, out int row, out _);
            return row;
        }
    }

    /// <summary>
    /// Gets the column index (1-based).
    /// </summary>
    public int Column
    {
        get
        {
            CellReference.Parse(Reference, out _, out int column);
            return column;
        }
    }

    // ── Value Methods ──────────────────────────────────────────────────

    /// <summary>
    /// Sets a string value.
    /// </summary>
    public void SetValue(string value)
    {
        if (value == null)
        {
            _cell.CellValue = null;
            _cell.DataType = null;
            return;
        }

        int index = _sharedStringManager.AddOrGetString(value);
        _cell.CellValue = new CellValue(index.ToString());
        _cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
    }

    /// <summary>
    /// Sets a numeric value.
    /// </summary>
    public void SetValue(double value)
    {
        _cell.CellValue = new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _cell.DataType = new EnumValue<CellValues>(CellValues.Number);
    }

    /// <summary>
    /// Sets an integer value.
    /// </summary>
    public void SetValue(int value) => SetValue((double)value);

    /// <summary>
    /// Sets a boolean value.
    /// </summary>
    public void SetValue(bool value)
    {
        _cell.CellValue = new CellValue(value ? "1" : "0");
        _cell.DataType = new EnumValue<CellValues>(CellValues.Boolean);
    }

    /// <summary>
    /// Sets a DateTime value.
    /// </summary>
    public void SetValue(DateTime value)
    {
        double oaDate = value.ToOADate();
        _cell.CellValue = new CellValue(oaDate.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _cell.DataType = new EnumValue<CellValues>(CellValues.Number);

        // Apply date format if no custom format is set
        if (_cell.StyleIndex == null || _cell.StyleIndex.Value == 0)
        {
            var style = new CellStyle();
            style.NumberFormat = ExcelNumberFormat.ShortDate;
            _cell.StyleIndex = _styleManager.GetOrCreateCellFormatIndex(style);
        }
    }

    /// <summary>
    /// Gets the cell value as an object.
    /// </summary>
    public object? GetValue()
    {
        if (_cell.CellValue == null)
            return null;

        string cellValue = _cell.CellValue.Text;

        if (_cell.DataType?.Value == CellValues.SharedString)
        {
            if (int.TryParse(cellValue, out int index))
                return _sharedStringManager.GetString(index);
            return cellValue;
        }

        if (_cell.DataType?.Value == CellValues.Boolean)
            return cellValue == "1";

        if (double.TryParse(cellValue, System.Globalization.NumberStyles.Any,
            System.Globalization.CultureInfo.InvariantCulture, out double numValue))
        {
            uint styleIdx = _cell.StyleIndex?.Value ?? 0;
            if (IsDateFormatIndex(styleIdx))
            {
                try { return DateTime.FromOADate(numValue); }
                catch { return numValue; }
            }
            return numValue;
        }

        return cellValue;
    }

    /// <summary>
    /// Gets the cell value as text.
    /// </summary>
    public string GetText() => GetValue()?.ToString() ?? string.Empty;

    // ── Formula Methods ────────────────────────────────────────────────

    /// <summary>
    /// Sets a formula.
    /// </summary>
    public void SetFormula(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
            throw new ArgumentException("Formula cannot be null or empty.", nameof(formula));

        formula = formula.TrimStart('=');
        _cell.CellFormula = new CellFormula(formula);
        _cell.CellValue = null;
    }

    /// <summary>
    /// Gets the formula.
    /// </summary>
    public string? GetFormula() => _cell.CellFormula?.Text;

    /// <summary>
    /// Checks if cell has a formula.
    /// </summary>
    public bool HasFormula => _cell.CellFormula != null;

    // ── Style Methods ──────────────────────────────────────────────────

    /// <summary>
    /// Gets or sets the cell style.
    /// </summary>
    public CellStyle Style
    {
        get => _styleManager.GetCellStyle(_cell.StyleIndex?.Value ?? 0);
        set
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            _cell.StyleIndex = _styleManager.GetOrCreateCellFormatIndex(value);
        }
    }

    /// <summary>
    /// Applies a style using a fluent builder action.
    /// Usage: cell.ApplyStyle(s => s.SetBold().SetFontColor(ExcelColor.Red));
    /// </summary>
    public Cell ApplyStyle(Action<CellStyle> styleAction)
    {
        if (styleAction == null) throw new ArgumentNullException(nameof(styleAction));

        // Get current style as base, then apply changes
        var style = _styleManager.GetCellStyle(_cell.StyleIndex?.Value ?? 0);
        styleAction(style);
        _cell.StyleIndex = _styleManager.GetOrCreateCellFormatIndex(style);
        return this;
    }

    /// <summary>
    /// Sets the font to bold.
    /// </summary>
    public Cell SetBold(bool bold = true)
    {
        ApplyStyle(s => s.SetBold(bold));
        return this;
    }

    /// <summary>
    /// Sets the font to italic.
    /// </summary>
    public Cell SetItalic(bool italic = true)
    {
        ApplyStyle(s => s.SetItalic(italic));
        return this;
    }

    /// <summary>
    /// Sets the font size.
    /// </summary>
    public Cell SetFontSize(double size)
    {
        ApplyStyle(s => s.SetFontSize(size));
        return this;
    }

    /// <summary>
    /// Sets the font color.
    /// </summary>
    public Cell SetFontColor(ExcelColor color)
    {
        ApplyStyle(s => s.SetFontColor(color));
        return this;
    }

    /// <summary>
    /// Sets the font name.
    /// </summary>
    public Cell SetFontName(string name)
    {
        ApplyStyle(s => s.SetFontName(name));
        return this;
    }

    /// <summary>
    /// Sets the background color (solid fill).
    /// </summary>
    public Cell SetBackgroundColor(ExcelColor color)
    {
        ApplyStyle(s => s.SetBackgroundColor(color));
        return this;
    }

    /// <summary>
    /// Sets the number format.
    /// </summary>
    public Cell SetNumberFormat(ExcelNumberFormat format)
    {
        ApplyStyle(s => s.SetNumberFormat(format));
        return this;
    }

    /// <summary>
    /// Sets the number format from a format code string.
    /// </summary>
    public Cell SetNumberFormat(string formatCode)
    {
        ApplyStyle(s => s.SetNumberFormat(formatCode));
        return this;
    }

    /// <summary>
    /// Sets the horizontal alignment.
    /// </summary>
    public Cell SetHorizontalAlignment(ExcelHorizontalAlignment alignment)
    {
        ApplyStyle(s => s.SetHorizontalAlignment(alignment));
        return this;
    }

    /// <summary>
    /// Sets the vertical alignment.
    /// </summary>
    public Cell SetVerticalAlignment(ExcelVerticalAlignment alignment)
    {
        ApplyStyle(s => s.SetVerticalAlignment(alignment));
        return this;
    }

    /// <summary>
    /// Sets text wrapping.
    /// </summary>
    public Cell SetWrapText(bool wrap = true)
    {
        ApplyStyle(s => s.SetWrapText(wrap));
        return this;
    }

    /// <summary>
    /// Sets all borders.
    /// </summary>
    public Cell SetAllBorders(ExcelBorderStyle style, ExcelColor? color = null)
    {
        ApplyStyle(s => s.SetAllBorders(style, color));
        return this;
    }

    // ── Clear ──────────────────────────────────────────────────────────

    /// <summary>
    /// Clears the cell content.
    /// </summary>
    public void Clear()
    {
        _cell.CellValue = null;
        _cell.DataType = null;
        _cell.CellFormula = null;
    }

    /// <summary>
    /// Clears the cell content and formatting.
    /// </summary>
    public void ClearAll()
    {
        Clear();
        _cell.StyleIndex = null;
    }

    // ── Internal helpers ───────────────────────────────────────────────

    /// <summary>
    /// Gets the underlying OpenXml cell.
    /// </summary>
    internal DocumentFormat.OpenXml.Spreadsheet.Cell OpenXmlCell => _cell;

    /// <summary>
    /// Checks if the given style index represents a date format.
    /// </summary>
    private bool IsDateFormatIndex(uint styleIndex)
    {
        if (styleIndex == 0) return false;

        var style = _styleManager.GetCellStyle(styleIndex);
        if (style.NumberFormat == null) return false;

        var fmtId = style.NumberFormat.FormatId;
        // Built-in date format IDs: 14-22, 45-47
        if ((fmtId >= 14 && fmtId <= 22) || (fmtId >= 45 && fmtId <= 47))
            return true;

        // Check format code for date patterns
        string code = style.NumberFormat.FormatCode.ToLowerInvariant();
        return code.Contains('y') || code.Contains('m') || code.Contains('d') ||
               code.Contains('h') || code.Contains("am") || code.Contains("pm");
    }
}
