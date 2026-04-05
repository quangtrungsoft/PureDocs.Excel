namespace TVE.PureDocs.Excel.Charts;

/// <summary>
/// Defines the type of chart to create.
/// Maps directly to OOXML chart element types (c:barChart, c:lineChart, etc.)
/// </summary>
public enum ExcelChartType
{
    // ── Column charts (vertical bars) ──────────────────────────────────────
    /// <summary>Clustered column chart (side-by-side vertical bars).</summary>
    ColumnClustered,

    /// <summary>Stacked column chart (stacked vertical bars).</summary>
    ColumnStacked,

    /// <summary>100% stacked column chart.</summary>
    ColumnStackedFull,

    // ── Bar charts (horizontal bars) ────────────────────────────────────────
    /// <summary>Clustered bar chart (side-by-side horizontal bars).</summary>
    BarClustered,

    /// <summary>Stacked bar chart (stacked horizontal bars).</summary>
    BarStacked,

    /// <summary>100% stacked bar chart.</summary>
    BarStackedFull,

    // ── Line charts ─────────────────────────────────────────────────────────
    /// <summary>Line chart with data points and lines.</summary>
    Line,

    /// <summary>Line chart with markers at data points.</summary>
    LineWithMarkers,

    /// <summary>Stacked line chart.</summary>
    LineStacked,

    /// <summary>Smooth line chart (no angular corners).</summary>
    LineSmooth,

    // ── Pie / Doughnut charts ───────────────────────────────────────────────
    /// <summary>Standard pie chart (single series only).</summary>
    Pie,

    /// <summary>Exploded pie chart.</summary>
    PieExploded,

    /// <summary>Doughnut chart.</summary>
    Doughnut,

    /// <summary>Exploded doughnut chart.</summary>
    DoughnutExploded,

    // ── Area charts ─────────────────────────────────────────────────────────
    /// <summary>Area chart (filled area under lines).</summary>
    Area,

    /// <summary>Stacked area chart.</summary>
    AreaStacked,

    /// <summary>100% stacked area chart.</summary>
    AreaStackedFull,

    // ── Scatter / Bubble charts ─────────────────────────────────────────────
    /// <summary>Scatter chart (XY data points, no lines).</summary>
    ScatterMarkers,

    /// <summary>Scatter chart with straight connecting lines.</summary>
    ScatterLines,

    /// <summary>Scatter chart with smooth connecting lines.</summary>
    ScatterSmooth,

    /// <summary>Bubble chart.</summary>
    Bubble,

    // ── Radar charts ────────────────────────────────────────────────────────
    /// <summary>Radar (spider web) chart.</summary>
    Radar,

    /// <summary>Radar chart filled with color.</summary>
    RadarFilled,

    // ── Stock charts ────────────────────────────────────────────────────────
    /// <summary>Stock (High-Low-Close) chart.</summary>
    Stock,
}

/// <summary>
/// Controls how multiple series are grouped in bar/column/area charts.
/// </summary>
public enum ChartGrouping
{
    /// <summary>Series displayed side-by-side.</summary>
    Clustered,

    /// <summary>Series stacked on top of each other.</summary>
    Stacked,

    /// <summary>Series stacked to reach 100% total.</summary>
    PercentStacked,

    /// <summary>Series drawn on top of each other (area charts).</summary>
    Standard
}

/// <summary>
/// Controls the position of the legend relative to the chart.
/// </summary>
public enum ChartLegendPosition
{
    /// <summary>Legend at the bottom of the chart.</summary>
    Bottom,

    /// <summary>Legend at the top of the chart.</summary>
    Top,

    /// <summary>Legend at the left side of the chart.</summary>
    Left,

    /// <summary>Legend at the right side of the chart.</summary>
    Right,

    /// <summary>Legend at the top-right corner.</summary>
    TopRight
}

/// <summary>
/// Horizontal alignment options for chart text elements.
/// </summary>
public enum ChartTextAlignment
{
    Left,
    Center,
    Right
}
