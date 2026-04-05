namespace TVE.PureDocs.Excel.Charts;

/// <summary>
/// Represents an Excel chart embedded in a worksheet.
/// Provides a fluent API for configuring chart type, series, axes, title, legend and position.
///
/// Usage pattern:
/// <code>
///   var chart = worksheet.AddChart(ExcelChartType.ColumnClustered, ChartPosition.From(2, 2, 16, 9))
///       .SetTitle("Monthly Revenue")
///       .AddSeries("Revenue", "Sheet1!$B$2:$B$13")
///           .SetCategories("Sheet1!$A$2:$A$13")
///       .ConfigureCategoryAxis(a => a.ShowGridlines())
///       .ConfigureValueAxis(a => a.SetNumberFormat("#,##0"))
///       .ConfigureLegend(l => l.SetPosition(ChartLegendPosition.Right));
///
///   worksheet.CommitChart(chart);
/// </code>
/// </summary>
public sealed class ExcelChart
{
    // ── Core properties ─────────────────────────────────────────────────────

    /// <summary>The chart type. Can be changed after construction.</summary>
    public ExcelChartType ChartType { get; private set; }

    /// <summary>Anchor position on the sheet.</summary>
    public ChartPosition Position { get; private set; }

    /// <summary>Chart title text. When null no title element is written.</summary>
    public string? Title { get; private set; }

    /// <summary>
    /// When true the chart has no title element (auto-deleted).
    /// Defaults to false (title element present but text may be empty).
    /// </summary>
    public bool AutoTitleDeleted { get; private set; }

    /// <summary>Whether a 3D chart variant is used.</summary>
    public bool Use3D { get; private set; }

    /// <summary>Data series. Ordered by Index (0-based).</summary>
    internal List<ChartSeries> Series { get; } = new();

    /// <summary>Category (X) axis options.</summary>
    public ChartAxisOptions CategoryAxis { get; } = new();

    /// <summary>Value (Y) axis options.</summary>
    public ChartAxisOptions ValueAxis { get; } = new();

    /// <summary>Legend options.</summary>
    public ChartLegendOptions Legend { get; } = new();

    /// <summary>
    /// Gap width between bar/column clusters (0–500). Default 150 (Excel default).
    /// Only applicable to bar/column charts.
    /// </summary>
    public int GapWidth { get; private set; } = 150;

    /// <summary>
    /// Overlap of bars within a cluster (-100 to 100). Default 0.
    /// Positive = overlapping, negative = gap between bars in cluster.
    /// </summary>
    public int Overlap { get; private set; } = 0;

    /// <summary>
    /// Whether data labels are shown on data points.
    /// </summary>
    public bool ShowDataLabels { get; private set; } = false;

    /// <summary>
    /// Hole size for doughnut charts (10–90, percent). Default 50.
    /// </summary>
    public int DoughnutHoleSize { get; private set; } = 50;

    /// <summary>
    /// Explosion percentage for exploded pie / doughnut charts (0–400).
    /// </summary>
    public int ExplosionPercent { get; private set; } = 0;

    /// <summary>Internal shape ID for the drawing element. Set by ChartWriter.</summary>
    internal int ShapeId { get; set; }

    // ── Constructor (internal — created via Worksheet.AddChart) ────────────

    internal ExcelChart(ExcelChartType chartType, ChartPosition position)
    {
        ChartType = chartType;
        Position = position ?? throw new ArgumentNullException(nameof(position));
    }

    // ── Fluent: Chart-level ─────────────────────────────────────────────────

    /// <summary>Sets the chart title.</summary>
    public ExcelChart SetTitle(string title)
    {
        Title = title;
        AutoTitleDeleted = false;
        return this;
    }

    /// <summary>Removes the chart title.</summary>
    public ExcelChart RemoveTitle()
    {
        Title = null;
        AutoTitleDeleted = true;
        return this;
    }

    /// <summary>Changes the chart type.</summary>
    public ExcelChart SetChartType(ExcelChartType type)
    {
        ChartType = type;
        return this;
    }

    /// <summary>Enables the 3D variant of the current chart type (where applicable).</summary>
    public ExcelChart Set3D(bool use3D = true)
    {
        Use3D = use3D;
        return this;
    }

    /// <summary>Sets the bar/column gap width (0–500). Default 150.</summary>
    public ExcelChart SetGapWidth(int gapWidth)
    {
        if (gapWidth < 0 || gapWidth > 500)
            throw new ArgumentOutOfRangeException(nameof(gapWidth), "Gap width must be 0–500.");
        GapWidth = gapWidth;
        return this;
    }

    /// <summary>Sets the bar/column overlap (-100 to 100). Default 0.</summary>
    public ExcelChart SetOverlap(int overlap)
    {
        if (overlap < -100 || overlap > 100)
            throw new ArgumentOutOfRangeException(nameof(overlap), "Overlap must be -100 to 100.");
        Overlap = overlap;
        return this;
    }

    /// <summary>Shows data labels on data points.</summary>
    public ExcelChart ShowLabels(bool show = true)
    {
        ShowDataLabels = show;
        return this;
    }

    /// <summary>
    /// Sets the doughnut hole size (10–90 percent).
    /// Only applies to <see cref="ExcelChartType.Doughnut"/> and
    /// <see cref="ExcelChartType.DoughnutExploded"/>.
    /// </summary>
    public ExcelChart SetDoughnutHoleSize(int percent)
    {
        if (percent < 10 || percent > 90)
            throw new ArgumentOutOfRangeException(nameof(percent), "Hole size must be 10–90.");
        DoughnutHoleSize = percent;
        return this;
    }

    /// <summary>Changes the anchor position.</summary>
    public ExcelChart SetPosition(ChartPosition position)
    {
        Position = position ?? throw new ArgumentNullException(nameof(position));
        return this;
    }

    // ── Fluent: Series ───────────────────────────────────────────────────────

    /// <summary>
    /// Adds a data series to the chart.
    /// </summary>
    /// <param name="name">
    /// Series legend name. Can be:
    /// • A plain label: "Revenue"
    /// • A cell reference formula: "Sheet1!$B$1"
    /// </param>
    /// <param name="valuesFormula">
    /// Range formula for the series values, e.g. "Sheet1!$B$2:$B$13".
    /// Leading '=' is stripped automatically.
    /// </param>
    /// <param name="nameIsFormula">
    /// Set to true when <paramref name="name"/> is a cell reference,
    /// false (default) when it is a plain string literal.
    /// </param>
    /// <returns>The created <see cref="ChartSeries"/> for further configuration.</returns>
    public ChartSeries AddSeries(string name, string valuesFormula, bool nameIsFormula = false)
    {
        var series = new ChartSeries(name, valuesFormula, nameIsFormula)
        {
            Index = Series.Count
        };
        Series.Add(series);
        return series;
    }

    /// <summary>Removes all series.</summary>
    public ExcelChart ClearSeries()
    {
        Series.Clear();
        return this;
    }

    // ── Fluent: Axes ────────────────────────────────────────────────────────

    /// <summary>
    /// Configures the category (X-axis) using an action on <see cref="ChartAxisOptions"/>.
    /// </summary>
    public ExcelChart ConfigureCategoryAxis(Action<ChartAxisOptions> configure)
    {
        configure?.Invoke(CategoryAxis);
        return this;
    }

    /// <summary>
    /// Configures the value (Y-axis) using an action on <see cref="ChartAxisOptions"/>.
    /// </summary>
    public ExcelChart ConfigureValueAxis(Action<ChartAxisOptions> configure)
    {
        configure?.Invoke(ValueAxis);
        return this;
    }

    // ── Fluent: Legend ──────────────────────────────────────────────────────

    /// <summary>
    /// Configures the legend using an action on <see cref="ChartLegendOptions"/>.
    /// </summary>
    public ExcelChart ConfigureLegend(Action<ChartLegendOptions> configure)
    {
        configure?.Invoke(Legend);
        return this;
    }

    // ── Internal helpers ────────────────────────────────────────────────────

    /// <summary>
    /// Determines the OOXML chart grouping value for the current chart type.
    /// </summary>
    internal string GetGroupingValue()
    {
        return ChartType switch
        {
            ExcelChartType.ColumnStacked or ExcelChartType.BarStacked or ExcelChartType.AreaStacked
                or ExcelChartType.LineStacked => "stacked",

            ExcelChartType.ColumnStackedFull or ExcelChartType.BarStackedFull
                or ExcelChartType.AreaStackedFull => "percentStacked",

            ExcelChartType.Area or ExcelChartType.Radar => "standard",

            _ => "clustered"
        };
    }

    /// <summary>
    /// Returns true when the chart type uses a bar element (horizontal bars).
    /// </summary>
    internal bool IsBarChart()
        => ChartType is ExcelChartType.BarClustered
            or ExcelChartType.BarStacked
            or ExcelChartType.BarStackedFull;

    /// <summary>
    /// Returns true for pie and doughnut charts (single-series, no axes).
    /// </summary>
    internal bool IsPieOrDoughnut()
        => ChartType is ExcelChartType.Pie or ExcelChartType.PieExploded
            or ExcelChartType.Doughnut or ExcelChartType.DoughnutExploded;

    /// <summary>
    /// Returns true for scatter/bubble charts (dual numeric axes).
    /// </summary>
    internal bool IsScatter()
        => ChartType is ExcelChartType.ScatterMarkers
            or ExcelChartType.ScatterLines
            or ExcelChartType.ScatterSmooth
            or ExcelChartType.Bubble;
}
