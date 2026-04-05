namespace TVE.PureDocs.Excel.Charts;

// ═══════════════════════════════════════════════════════════════════════════
// ChartAxisOptions
// ═══════════════════════════════════════════════════════════════════════════

/// <summary>
/// Configuration for a single chart axis (category or value).
/// All properties are optional — Excel defaults apply when null.
/// </summary>
public sealed class ChartAxisOptions
{
    /// <summary>Axis title text. When null no title is shown.</summary>
    public string? Title { get; set; }

    /// <summary>Whether the axis is visible. Defaults to true.</summary>
    public bool Visible { get; set; } = true;

    /// <summary>Whether major gridlines are shown. Defaults to false.</summary>
    public bool ShowMajorGridlines { get; set; } = false;

    /// <summary>Whether minor gridlines are shown. Defaults to false.</summary>
    public bool ShowMinorGridlines { get; set; } = false;

    /// <summary>
    /// Number format code applied to value axis tick labels.
    /// e.g., "#,##0", "0.0%", "$#,##0.00"
    /// </summary>
    public string? NumberFormat { get; set; }

    /// <summary>Minimum value of the axis scale. When null Excel auto-scales.</summary>
    public double? MinValue { get; set; }

    /// <summary>Maximum value of the axis scale. When null Excel auto-scales.</summary>
    public double? MaxValue { get; set; }

    /// <summary>Major unit interval. When null Excel auto-calculates.</summary>
    public double? MajorUnit { get; set; }

    /// <summary>
    /// Text rotation in degrees for axis tick labels (0–360).
    /// Typical usage: 45 for diagonal, 90 for vertical.
    /// </summary>
    public int? LabelRotation { get; set; }

    /// <summary>Font size in points for tick labels.</summary>
    public double? FontSize { get; set; }

    /// <summary>Major tick mark style. Default is Outside.</summary>
    public ChartTickMark MajorTickMark { get; set; } = ChartTickMark.Outside;

    /// <summary>Minor tick mark style. Default is None.</summary>
    public ChartTickMark MinorTickMark { get; set; } = ChartTickMark.None;

    // ── Fluent API ──────────────────────────────────────────────────────────

    public ChartAxisOptions SetTitle(string title) { Title = title; return this; }
    public ChartAxisOptions Hide() { Visible = false; return this; }
    public ChartAxisOptions ShowGridlines(bool major = true, bool minor = false)
    {
        ShowMajorGridlines = major;
        ShowMinorGridlines = minor;
        return this;
    }
    public ChartAxisOptions SetNumberFormat(string format) { NumberFormat = format; return this; }
    public ChartAxisOptions SetScale(double? min = null, double? max = null, double? majorUnit = null)
    {
        MinValue = min;
        MaxValue = max;
        MajorUnit = majorUnit;
        return this;
    }
    public ChartAxisOptions SetLabelRotation(int degrees) { LabelRotation = degrees; return this; }
    public ChartAxisOptions SetFontSize(double size) { FontSize = size; return this; }
    
    /// <summary>Sets major and minor tick mark styles.</summary>
    public ChartAxisOptions SetTickMarks(ChartTickMark major = ChartTickMark.Outside, ChartTickMark minor = ChartTickMark.None)
    {
        MajorTickMark = major;
        MinorTickMark = minor;
        return this;
    }
}

/// <summary>
/// Tick mark style for chart axes.
/// </summary>
public enum ChartTickMark
{
    /// <summary>No tick marks.</summary>
    None,
    /// <summary>Tick marks inside the plot area.</summary>
    Inside,
    /// <summary>Tick marks outside the plot area (default for major).</summary>
    Outside,
    /// <summary>Tick marks crossing the axis line.</summary>
    Cross
}

// ═══════════════════════════════════════════════════════════════════════════
// ChartLegendOptions
// ═══════════════════════════════════════════════════════════════════════════

/// <summary>
/// Configuration for the chart legend.
/// </summary>
public sealed class ChartLegendOptions
{
    /// <summary>Whether the legend is shown. Default true.</summary>
    public bool Visible { get; set; } = true;

    /// <summary>Position of the legend relative to the plot area.</summary>
    public ChartLegendPosition Position { get; set; } = ChartLegendPosition.Bottom;

    /// <summary>
    /// When true the legend does NOT overlay the plot area.
    /// This is the Excel default (overlay = false).
    /// </summary>
    public bool OverlayChart { get; set; } = false;

    /// <summary>Font size of legend text in points.</summary>
    public double? FontSize { get; set; }

    // ── Fluent API ──────────────────────────────────────────────────────────

    public ChartLegendOptions Hide() { Visible = false; return this; }
    public ChartLegendOptions SetPosition(ChartLegendPosition pos) { Position = pos; return this; }
    public ChartLegendOptions SetFontSize(double size) { FontSize = size; return this; }
    public ChartLegendOptions SetOverlay(bool overlay) { OverlayChart = overlay; return this; }
}

// ═══════════════════════════════════════════════════════════════════════════
// ChartPosition
// ═══════════════════════════════════════════════════════════════════════════

/// <summary>
/// Defines the on-sheet position and size of a chart using a two-cell anchor.
/// In OOXML this maps to xdr:twoCellAnchor → xdr:from / xdr:to.
/// Coordinates are zero-based column/row indices (unlike 1-based Cell API).
/// </summary>
public sealed class ChartPosition
{
    /// <summary>Zero-based column index of the top-left anchor cell.</summary>
    public int FromColumn { get; set; }

    /// <summary>Zero-based row index of the top-left anchor cell.</summary>
    public int FromRow { get; set; }

    /// <summary>Column offset in EMU from the top-left anchor cell edge. Default 0.</summary>
    public long FromColumnOffset { get; set; } = 0;

    /// <summary>Row offset in EMU from the top-left anchor cell edge. Default 0.</summary>
    public long FromRowOffset { get; set; } = 0;

    /// <summary>Zero-based column index of the bottom-right anchor cell (exclusive).</summary>
    public int ToColumn { get; set; }

    /// <summary>Zero-based row index of the bottom-right anchor cell (exclusive).</summary>
    public int ToRow { get; set; }

    /// <summary>Column offset in EMU from the bottom-right anchor cell edge. Default 0.</summary>
    public long ToColumnOffset { get; set; } = 0;

    /// <summary>Row offset in EMU from the bottom-right anchor cell edge. Default 0.</summary>
    public long ToRowOffset { get; set; } = 0;

    // ── Factory helpers ─────────────────────────────────────────────────────

    /// <summary>
    /// Creates a chart position from 1-based cell references (matches Cell API).
    /// Example: ChartPosition.From(row:2, col:2, toRow:16, toCol:9)
    /// → chart occupying B2:I16.
    /// </summary>
    /// <param name="row">1-based top-left row (1 to 1048576).</param>
    /// <param name="col">1-based top-left column (1 to 16384).</param>
    /// <param name="toRow">1-based bottom-right row (exclusive — next row after chart).</param>
    /// <param name="toCol">1-based bottom-right column (exclusive).</param>
    public static ChartPosition From(int row, int col, int toRow, int toCol)
    {
        // Excel limits: max rows = 1048576, max columns = 16384
        const int MaxRows = 1048576;
        const int MaxCols = 16384;
        
        if (row < 1 || row > MaxRows)
            throw new ArgumentOutOfRangeException(nameof(row), $"Row must be between 1 and {MaxRows}.");
        if (col < 1 || col > MaxCols)
            throw new ArgumentOutOfRangeException(nameof(col), $"Column must be between 1 and {MaxCols}.");
        if (toRow < row || toRow > MaxRows + 1)
            throw new ArgumentOutOfRangeException(nameof(toRow), $"toRow must be between {row} and {MaxRows + 1}.");
        if (toCol < col || toCol > MaxCols + 1)
            throw new ArgumentOutOfRangeException(nameof(toCol), $"toCol must be between {col} and {MaxCols + 1}.");

        // Convert 1-based → 0-based for OOXML twoCellAnchor
        return new ChartPosition
        {
            FromRow    = row - 1,
            FromColumn = col - 1,
            ToRow      = toRow - 1,
            ToColumn   = toCol - 1
        };
    }

    /// <summary>
    /// Creates a chart position with exact 0-based indices (raw OOXML coordinates).
    /// </summary>
    public static ChartPosition FromZeroBased(int fromRow, int fromCol, int toRow, int toCol)
        => new ChartPosition
        {
            FromRow    = fromRow,
            FromColumn = fromCol,
            ToRow      = toRow,
            ToColumn   = toCol
        };
}
