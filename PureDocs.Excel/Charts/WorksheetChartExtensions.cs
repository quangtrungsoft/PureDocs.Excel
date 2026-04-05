using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Packaging;
using TVE.PureDocs.Excel.Charts;

namespace TVE.PureDocs.Excel;

/// <summary>
/// Extends <see cref="Worksheet"/> with chart creation capabilities.
///
/// Two-step pattern (mirrors how Excel internally separates configuration from serialization):
/// <code>
/// // Step 1: Configure
/// var chart = worksheet.AddChart(ExcelChartType.ColumnClustered,
///                                ChartPosition.From(row: 2, col: 2, toRow: 18, toCol: 9))
///     .SetTitle("Monthly Revenue")
///     .AddSeries("Revenue", "Sheet1!$B$2:$B$13")
///         .SetCategories("Sheet1!$A$2:$A$13")
///     .ConfigureValueAxis(a => a.SetNumberFormat("#,##0").ShowGridlines())
///     .ConfigureLegend(l => l.SetPosition(ChartLegendPosition.Bottom));
///
/// // Step 2: Commit (write OOXML)
/// worksheet.CommitChart(chart);
/// </code>
/// </summary>
public static class WorksheetChartExtensions
{
    // Per-worksheet ChartWriter instance (one drawing part per worksheet)
    // Stored via ConditionalWeakTable to avoid modifying sealed Worksheet class.
    private static readonly ConditionalWeakTable<Worksheet, ChartWriter> _writerTable = new();

    /// <summary>
    /// Creates a new <see cref="ExcelChart"/> configuration object anchored to this worksheet.
    ///
    /// DOES NOT write any XML yet — call <see cref="CommitChart"/> to serialize.
    /// This separation allows you to inspect / test the chart configuration before committing.
    /// </summary>
    /// <param name="worksheet">The worksheet that will contain the chart.</param>
    /// <param name="chartType">Chart type (column, bar, line, pie, …).</param>
    /// <param name="position">
    /// Anchor position. Use <see cref="ChartPosition.From(int,int,int,int)"/> for
    /// 1-based row/column convenience.
    /// </param>
    public static ExcelChart AddChart(
        this Worksheet worksheet,
        ExcelChartType chartType,
        ChartPosition position)
    {
        if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
        if (position  == null) throw new ArgumentNullException(nameof(position));

        return new ExcelChart(chartType, position);
    }

    /// <summary>
    /// Overload: creates a chart with a default position spanning B2:I18.
    /// Useful for quick prototyping.
    /// </summary>
    public static ExcelChart AddChart(
        this Worksheet worksheet,
        ExcelChartType chartType)
        => worksheet.AddChart(chartType, ChartPosition.From(2, 2, 18, 9));

    /// <summary>
    /// Serializes the <see cref="ExcelChart"/> configuration into the worksheet's OOXML.
    ///
    /// After this call:
    /// - A <see cref="DrawingsPart"/> is created (or reused) on the worksheet.
    /// - A <see cref="ChartPart"/> containing c:chartSpace XML is added.
    /// - A <c>xdr:twoCellAnchor</c> element is written into the drawing part.
    /// - A <c>&lt;drawing r:id="…"/&gt;</c> reference is appended to the worksheet element.
    ///
    /// Multiple charts can be committed to the same worksheet — each call adds a new anchor.
    /// </summary>
    /// <param name="worksheet">The worksheet that owns the chart.</param>
    /// <param name="chart">The chart configuration to serialize.</param>
    /// <returns>
    /// The <see cref="ExcelChart"/> instance (for optional method chaining after commit).
    /// </returns>
    /// <exception cref="InvalidOperationException">
    /// Thrown when the chart has no series defined or series have invalid formulas.
    /// </exception>
    /// <exception cref="ArgumentException">
    /// Thrown when chart position is invalid.
    /// </exception>
    public static ExcelChart CommitChart(this Worksheet worksheet, ExcelChart chart)
    {
        if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
        if (chart == null)     throw new ArgumentNullException(nameof(chart));

        // Validate series count
        if (chart.Series.Count == 0)
            throw new InvalidOperationException(
                "Chart must have at least one series. Call chart.AddSeries() before CommitChart().");

        // Validate each series has a valid values formula
        foreach (var series in chart.Series)
        {
            if (string.IsNullOrWhiteSpace(series.ValuesFormula))
                throw new InvalidOperationException(
                    $"Series '{series.Name}' has empty ValuesFormula. Each series must have a data range.");
        }

        // Validate position is within reasonable bounds
        if (chart.Position.FromRow < 0 || chart.Position.FromColumn < 0)
            throw new ArgumentException(
                "Chart position cannot have negative row or column values.", nameof(chart));
        
        if (chart.Position.ToRow < chart.Position.FromRow || chart.Position.ToColumn < chart.Position.FromColumn)
            throw new ArgumentException(
                "Chart 'To' position must be greater than or equal to 'From' position.", nameof(chart));

        var worksheetPart = worksheet.GetWorksheetPart();
        var writer = _writerTable.GetOrCreateValue(worksheet);
        writer.WriteChart(worksheetPart, chart);

        return chart;
    }

    /// <summary>
    /// Convenience: combines <see cref="AddChart(Worksheet,ExcelChartType,ChartPosition)"/>
    /// and <see cref="CommitChart"/> in a single call.
    ///
    /// Use when you don't need to inspect the chart object before serialization.
    /// </summary>
    /// <param name="configure">
    /// Action to configure the chart (add series, set title, configure axes, etc.)
    /// </param>
    public static ExcelChart AddAndCommitChart(
        this Worksheet worksheet,
        ExcelChartType chartType,
        ChartPosition position,
        Action<ExcelChart> configure)
    {
        if (configure == null) throw new ArgumentNullException(nameof(configure));

        var chart = worksheet.AddChart(chartType, position);
        configure(chart);
        return worksheet.CommitChart(chart);
    }
}
