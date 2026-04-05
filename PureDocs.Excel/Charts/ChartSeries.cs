namespace TVE.PureDocs.Excel.Charts;

/// <summary>
/// Represents a single data series in a chart.
/// A series binds:
///   - Name  → displayed in the legend and tooltip
///   - Values → the numeric data range (required)
///   - Categories → the label range for the X-axis (optional, shared across series)
/// </summary>
public sealed class ChartSeries
{
    /// <summary>
    /// Display name of the series.
    /// Can be a plain string literal ("Revenue") or an Excel range formula ("Sheet1!$B$1").
    /// </summary>
    public string Name { get; private set; }

    /// <summary>
    /// Whether <see cref="Name"/> is a cell formula reference (true) or a literal string (false).
    /// </summary>
    public bool NameIsFormula { get; private set; }

    /// <summary>
    /// Excel range formula for the series values (e.g., "Sheet1!$B$2:$B$10").
    /// Required — must be a valid range formula.
    /// </summary>
    public string ValuesFormula { get; private set; }

    /// <summary>
    /// Excel range formula for the category labels (e.g., "Sheet1!$A$2:$A$10").
    /// Optional — when null, Excel uses 1, 2, 3 … as default categories.
    /// </summary>
    public string? CategoriesFormula { get; private set; }

    /// <summary>
    /// Optional solid fill color override for this series (hex ARGB, e.g. "FF4472C4").
    /// When null the default Excel color cycle is used.
    /// </summary>
    public string? FillColorHex { get; private set; }

    /// <summary>
    /// Optional marker style for line / scatter charts.
    /// </summary>
    public ChartMarkerStyle MarkerStyle { get; private set; } = ChartMarkerStyle.Automatic;

    /// <summary>
    /// Optional explicit series order index.
    /// Automatically assigned by <see cref="ExcelChart"/> if not set.
    /// </summary>
    internal int Index { get; set; }

    // ── Constructor (internal — use ExcelChart.AddSeries) ──────────────────

    internal ChartSeries(string name, string valuesFormula, bool nameIsFormula = false)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Series name cannot be null or empty.", nameof(name));
        if (string.IsNullOrWhiteSpace(valuesFormula))
            throw new ArgumentException("Values formula cannot be null or empty.", nameof(valuesFormula));

        Name = name;
        ValuesFormula = NormalizeFormula(valuesFormula);
        NameIsFormula = nameIsFormula;
    }

    // ── Fluent API ──────────────────────────────────────────────────────────

    /// <summary>
    /// Sets the category (X-axis label) range formula.
    /// </summary>
    /// <param name="formula">Range formula, e.g. "Sheet1!$A$2:$A$10".</param>
    public ChartSeries SetCategories(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
            throw new ArgumentException("Categories formula cannot be null or empty.", nameof(formula));

        CategoriesFormula = NormalizeFormula(formula);
        return this;
    }

    /// <summary>
    /// Sets the fill color for this series using a hex color string.
    /// </summary>
    /// <param name="hexColor">
    /// RGB hex string (e.g., "4472C4" or "#4472C4").
    /// ARGB is also accepted (e.g., "FF4472C4") but note that alpha channel
    /// is not supported for series fill colors and will be ignored.
    /// </param>
    /// <exception cref="ArgumentException">Thrown if hexColor is not a valid hex color.</exception>
    public ChartSeries SetFillColor(string hexColor)
    {
        if (string.IsNullOrWhiteSpace(hexColor))
            throw new ArgumentException("Color cannot be null or empty.", nameof(hexColor));

        hexColor = hexColor.TrimStart('#').ToUpperInvariant();

        // Validate hex format: must be 6 (RGB) or 8 (ARGB) hex digits
        if (hexColor.Length != 6 && hexColor.Length != 8)
            throw new ArgumentException($"Invalid hex color format: '{hexColor}'. Expected 6 (RGB) or 8 (ARGB) hex digits.", nameof(hexColor));

        // Validate all characters are hex digits
        foreach (char c in hexColor)
        {
            if (!((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F')))
                throw new ArgumentException($"Invalid hex color: '{hexColor}'. Contains non-hex character '{c}'.", nameof(hexColor));
        }

        // Normalize to 8-char ARGB (prepend FF for full opacity if 6-char RGB)
        FillColorHex = hexColor.Length == 6 ? "FF" + hexColor : hexColor;
        return this;
    }

    /// <summary>
    /// Sets the marker style for line and scatter charts.
    /// </summary>
    public ChartSeries SetMarkerStyle(ChartMarkerStyle style)
    {
        MarkerStyle = style;
        return this;
    }

    // ── Helpers ─────────────────────────────────────────────────────────────

    // Basic pattern for range formulas: SheetName!A1:B10 or SheetName!$A$1:$B$10
    // Also handles quoted sheet names: 'Sheet Name'!A1:B10
    private static readonly System.Text.RegularExpressions.Regex RangeFormulaPattern = 
        new System.Text.RegularExpressions.Regex(
            @"^('?[^'!]+'?|[^!]+)!\$?[A-Z]+\$?\d+(:\$?[A-Z]+\$?\d+)?$",
            System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.IgnoreCase);

    /// <summary>
    /// Normalizes and validates a formula string.
    /// </summary>
    /// <param name="formula">The formula to normalize.</param>
    /// <param name="validate">If true, validates the formula format (default true).</param>
    /// <returns>Normalized formula string.</returns>
    /// <exception cref="ArgumentException">Thrown if formula format is invalid and validate is true.</exception>
    private static string NormalizeFormula(string formula, bool validate = true)
    {
        string normalized = formula.TrimStart('=');
        
        if (validate && !RangeFormulaPattern.IsMatch(normalized))
        {
            throw new ArgumentException(
                $"Invalid range formula format: '{formula}'. Expected format: 'SheetName!$A$1:$B$10' or 'SheetName!A1:B10'.",
                nameof(formula));
        }
        
        return normalized;
    }
}

/// <summary>
/// Marker styles for line and scatter chart data points.
/// Maps to OOXML c:marker/c:symbol@val values.
/// </summary>
public enum ChartMarkerStyle
{
    /// <summary>Excel chooses the marker automatically.</summary>
    Automatic,
    /// <summary>No marker shown.</summary>
    None,
    Circle,
    Square,
    Diamond,
    Triangle,
    Star,
    Plus,
    X
}
