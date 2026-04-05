using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

// Alias sprawling OpenXml namespaces for readability
using C  = DocumentFormat.OpenXml.Drawing.Charts;
using A  = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace TVE.PureDocs.Excel.Charts;

/// <summary>
/// Internal engine that translates <see cref="ExcelChart"/> configuration into OOXML.
///
/// Responsibilities:
///   1. Create / reuse the <see cref="DrawingsPart"/> on the worksheet.
///   2. Write the <c>xdr:twoCellAnchor</c> drawing element.
///   3. Create a <see cref="ChartPart"/> and write the full chart XML tree:
///         c:chartSpace → c:chart → c:plotArea → c:{barChart|lineChart|…}
///   4. Handle all supported chart types via a dispatch table.
///   5. Write axes (catAx / valAx) and legend.
///
/// Design mirrors <see cref="StyleManager"/>: pure OpenXML manipulation, no
/// public surface area, injected only by <see cref="Worksheet.CommitChart"/>.
/// </summary>
internal sealed class ChartWriter
{
    // Shape ID counter per worksheet (xdr:graphicFrame requires unique @id)
    // Uses Interlocked for thread-safety in case of concurrent chart writes
    private int _nextShapeId = 0;

    // OpenXML axis ID constants (must be consistent within one chart)
    private const uint CatAxisId = 2093507336;
    private const uint ValAxisId = 2093510856;

    // ── Namespace URI Constants ─────────────────────────────────────────────
    // Cached for performance - avoids repeated string allocations
    private const string NsChart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private const string NsDrawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string NsRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string NsSpreadsheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    // ── Entry point ─────────────────────────────────────────────────────────

    /// <summary>
    /// Writes the chart into the worksheet and returns the <see cref="ChartPart"/>
    /// relationship ID (rId) for diagnostics / test assertions.
    /// </summary>
    public string WriteChart(WorksheetPart worksheetPart, ExcelChart chart)
    {
        if (worksheetPart == null) throw new ArgumentNullException(nameof(worksheetPart));
        if (chart == null)         throw new ArgumentNullException(nameof(chart));

        // 1 ── Ensure DrawingsPart exists on the worksheet
        var drawingsPart = EnsureDrawingsPart(worksheetPart);

        // 2 ── Create a ChartPart inside the DrawingsPart
        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        string chartRelId = drawingsPart.GetIdOfPart(chartPart);

        // 3 ── Write the chart XML (c:chartSpace)
        chartPart.ChartSpace = BuildChartSpace(chart);
        chartPart.ChartSpace.Save();

        // 4 ── Add twoCellAnchor to the worksheet drawing
        // Thread-safe increment for concurrent chart writes
        int shapeId = System.Threading.Interlocked.Increment(ref _nextShapeId);
        chart.ShapeId = shapeId;
        var anchor = BuildTwoCellAnchor(chart.Position, chartRelId, shapeId);

        var drawing = GetOrCreateDrawing(drawingsPart);
        drawing.Append(anchor);
        drawingsPart.WorksheetDrawing.Save();

        return chartRelId;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // DrawingsPart management
    // ═══════════════════════════════════════════════════════════════════════

    private static DrawingsPart EnsureDrawingsPart(WorksheetPart worksheetPart)
    {
        if (worksheetPart.DrawingsPart != null)
            return worksheetPart.DrawingsPart;

        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

        // Initialize empty WorksheetDrawing
        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
        drawingsPart.WorksheetDrawing.AddNamespaceDeclaration("xdr", NsSpreadsheetDrawing);
        drawingsPart.WorksheetDrawing.AddNamespaceDeclaration("a", NsDrawing);
        drawingsPart.WorksheetDrawing.Save();

        // Register the drawing reference on the worksheet element
        string drawingRelId = worksheetPart.GetIdOfPart(drawingsPart);
        var wsElement = worksheetPart.Worksheet;

        // Check if a <drawing> element already exists (e.g., from a loaded file or manual addition)
        var existingDrawingRef = wsElement.Elements<DocumentFormat.OpenXml.Spreadsheet.Drawing>()
            .FirstOrDefault();
        
        if (existingDrawingRef == null)
        {
            // <drawing r:id="rId1"/> must be appended after <sheetData>
            var drawingRef = new DocumentFormat.OpenXml.Spreadsheet.Drawing
            {
                Id = drawingRelId
            };

            // Append after the last known child (SheetData, MergeCells, AutoFilter…)
            wsElement.Append(drawingRef);
            wsElement.Save();
        }
        else if (existingDrawingRef.Id != drawingRelId)
        {
            // Update existing reference to point to new DrawingsPart
            existingDrawingRef.Id = drawingRelId;
            wsElement.Save();
        }

        return drawingsPart;
    }

    private static Xdr.WorksheetDrawing GetOrCreateDrawing(DrawingsPart drawingsPart)
    {
        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();
        return drawingsPart.WorksheetDrawing;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // TwoCellAnchor — positions the chart frame on the sheet grid
    // ═══════════════════════════════════════════════════════════════════════

    private static Xdr.TwoCellAnchor BuildTwoCellAnchor(
        ChartPosition pos, string chartRelId, int shapeId)
    {
        var anchor = new Xdr.TwoCellAnchor();

        // From cell
        anchor.Append(new Xdr.FromMarker(
            new Xdr.ColumnId(pos.FromColumn.ToString()),
            new Xdr.ColumnOffset(pos.FromColumnOffset.ToString()),
            new Xdr.RowId(pos.FromRow.ToString()),
            new Xdr.RowOffset(pos.FromRowOffset.ToString())));

        // To cell
        anchor.Append(new Xdr.ToMarker(
            new Xdr.ColumnId(pos.ToColumn.ToString()),
            new Xdr.ColumnOffset(pos.ToColumnOffset.ToString()),
            new Xdr.RowId(pos.ToRow.ToString()),
            new Xdr.RowOffset(pos.ToRowOffset.ToString())));

        // GraphicFrame
        var graphicFrame = new Xdr.GraphicFrame { Macro = "" };

        // Non-visual frame properties
        graphicFrame.Append(new Xdr.NonVisualGraphicFrameProperties(
            new Xdr.NonVisualDrawingProperties
            {
                Id = (uint)shapeId,
                Name = $"Chart {shapeId}"
            },
            new Xdr.NonVisualGraphicFrameDrawingProperties()));

        // Transform (required even if zeroed)
        graphicFrame.Append(new Xdr.Transform(
            new A.Offset { X = 0, Y = 0 },
            new A.Extents { Cx = 0, Cy = 0 }));

        // Graphic → graphicData → c:chart reference
        var graphic = new A.Graphic(
            new A.GraphicData(
                new C.ChartReference { Id = chartRelId })
            { Uri = NsChart });

        graphicFrame.Append(graphic);
        anchor.Append(graphicFrame);
        anchor.Append(new Xdr.ClientData()); // required by schema

        return anchor;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // ChartSpace root
    // ═══════════════════════════════════════════════════════════════════════

    private static C.ChartSpace BuildChartSpace(ExcelChart chart)
    {
        var chartSpace = new C.ChartSpace();
        chartSpace.AddNamespaceDeclaration("c", NsChart);
        chartSpace.AddNamespaceDeclaration("a", NsDrawing);
        chartSpace.AddNamespaceDeclaration("r", NsRelationships);

        // Language
        chartSpace.Append(new EditingLanguage { Val = "en-US" });
        
        // Rounded corners - Excel default is false (square corners)
        chartSpace.Append(new C.RoundedCorners { Val = false });

        // Build inner c:chart element
        var chartElement = new C.Chart();

        // Title
        if (!chart.AutoTitleDeleted && !string.IsNullOrEmpty(chart.Title))
            chartElement.Append(BuildTitle(chart.Title!));
        else if (chart.AutoTitleDeleted)
            chartElement.Append(new C.AutoTitleDeleted { Val = true });

        // Plot area
        chartElement.Append(BuildPlotArea(chart));

        // Legend
        if (chart.Legend.Visible)
            chartElement.Append(BuildLegend(chart.Legend));

        // Plot visible only
        chartElement.Append(new C.PlotVisibleOnly { Val = true });

        chartSpace.Append(chartElement);
        return chartSpace;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Title
    // ═══════════════════════════════════════════════════════════════════════

    private static C.Title BuildTitle(string titleText)
    {
        var title = new C.Title();

        var txPr = new C.ChartText(
            new C.RichText(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.Run(
                        new A.Text(titleText)))));

        title.Append(txPr);
        title.Append(new C.Overlay { Val = false });
        return title;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // PlotArea + chart-type dispatch
    // ═══════════════════════════════════════════════════════════════════════

    private static C.PlotArea BuildPlotArea(ExcelChart chart)
    {
        var plotArea = new C.PlotArea();
        plotArea.Append(new C.Layout()); // auto layout

        // Dispatch to the appropriate chart element builder
        OpenXmlCompositeElement chartElement = chart.ChartType switch
        {
            ExcelChartType.ColumnClustered or ExcelChartType.ColumnStacked
                or ExcelChartType.ColumnStackedFull
                => BuildBarColumnChart(chart, isBar: false),

            ExcelChartType.BarClustered or ExcelChartType.BarStacked
                or ExcelChartType.BarStackedFull
                => BuildBarColumnChart(chart, isBar: true),

            ExcelChartType.Line or ExcelChartType.LineWithMarkers
                or ExcelChartType.LineStacked or ExcelChartType.LineSmooth
                => BuildLineChart(chart),

            ExcelChartType.Pie or ExcelChartType.PieExploded
                => BuildPieChart(chart),

            ExcelChartType.Doughnut or ExcelChartType.DoughnutExploded
                => BuildDoughnutChart(chart),

            ExcelChartType.Area or ExcelChartType.AreaStacked or ExcelChartType.AreaStackedFull
                => BuildAreaChart(chart),

            ExcelChartType.ScatterMarkers or ExcelChartType.ScatterLines or ExcelChartType.ScatterSmooth
                => BuildScatterChart(chart),

            ExcelChartType.Radar or ExcelChartType.RadarFilled
                => BuildRadarChart(chart),

            ExcelChartType.Stock
                => throw new NotSupportedException(
                    "Stock charts require special High-Low-Close series structure and are not yet implemented. " +
                    "Consider using a Line chart with multiple series as an alternative."),

            _ => BuildBarColumnChart(chart, isBar: false) // safe default
        };

        plotArea.Append(chartElement);

        // Axes — not added for pie/doughnut/radar (they use different axis models)
        if (!chart.IsPieOrDoughnut())
        {
            plotArea.Append(BuildCategoryAxis(chart.CategoryAxis, chart));
            plotArea.Append(BuildValueAxis(chart.ValueAxis, chart));
        }

        return plotArea;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Bar / Column chart  (c:barChart)
    // ═══════════════════════════════════════════════════════════════════════

    private static C.BarChart BuildBarColumnChart(ExcelChart chart, bool isBar)
    {
        var barChart = new C.BarChart();

        barChart.Append(new C.BarDirection
        {
            Val = isBar
                ? C.BarDirectionValues.Bar   // horizontal
                : C.BarDirectionValues.Column // vertical
        });

        barChart.Append(new C.BarGrouping { Val = GroupingToBarGrouping(chart.GetGroupingValue()) });

        // Overlap and gap
        barChart.Append(new C.Overlap { Val = (sbyte)chart.Overlap });
        barChart.Append(new C.GapWidth { Val = (ushort)chart.GapWidth });

        // Series
        foreach (var series in chart.Series)
            barChart.Append(BuildBarColumnSeries(series, chart));

        // Data labels
        barChart.Append(BuildDataLabels(chart.ShowDataLabels));

        // Axis references
        barChart.Append(new C.AxisId { Val = CatAxisId });
        barChart.Append(new C.AxisId { Val = ValAxisId });

        return barChart;
    }

    private static C.BarChartSeries BuildBarColumnSeries(ChartSeries series, ExcelChart chart)
    {
        var ser = new C.BarChartSeries();
        ser.Append(new C.Index { Val = (uint)series.Index });
        ser.Append(new C.Order { Val = (uint)series.Index });

        // Series name
        ser.Append(BuildSeriesName(series));

        // Fill color override
        if (series.FillColorHex != null)
            ser.Append(BuildSolidFillShapeProperties(series.FillColorHex));

        // Categories
        if (series.CategoriesFormula != null)
            ser.Append(BuildCategoryAxisData(series.CategoriesFormula));

        // Values
        ser.Append(BuildNumericValues(series.ValuesFormula));

        return ser;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Line chart  (c:lineChart)
    // ═══════════════════════════════════════════════════════════════════════

    private static C.LineChart BuildLineChart(ExcelChart chart)
    {
        var lineChart = new C.LineChart();

        lineChart.Append(new C.Grouping { Val = GroupingToGroupingValues(chart.GetGroupingValue()) });

        foreach (var series in chart.Series)
            lineChart.Append(BuildLineChartSeries(series, chart));

        lineChart.Append(BuildDataLabels(chart.ShowDataLabels));
        lineChart.Append(new C.AxisId { Val = CatAxisId });
        lineChart.Append(new C.AxisId { Val = ValAxisId });

        return lineChart;
    }

    private static C.LineChartSeries BuildLineChartSeries(ChartSeries series, ExcelChart chart)
    {
        var ser = new C.LineChartSeries();
        ser.Append(new C.Index { Val = (uint)series.Index });
        ser.Append(new C.Order { Val = (uint)series.Index });
        ser.Append(BuildSeriesName(series));

        bool showMarkers = chart.ChartType is ExcelChartType.LineWithMarkers;
        bool smooth = chart.ChartType is ExcelChartType.LineSmooth;

        // Marker
        var marker = new C.Marker();
        marker.Append(new C.Symbol
        {
            Val = showMarkers
                ? MarkerStyleToOxml(series.MarkerStyle)
                : C.MarkerStyleValues.None
        });
        ser.Append(marker);

        if (series.CategoriesFormula != null)
            ser.Append(BuildCategoryAxisData(series.CategoriesFormula));

        ser.Append(BuildNumericValues(series.ValuesFormula));
        ser.Append(new C.Smooth { Val = smooth });

        return ser;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Pie chart  (c:pieChart)
    // ═══════════════════════════════════════════════════════════════════════

    private static C.PieChart BuildPieChart(ExcelChart chart)
    {
        var pieChart = new C.PieChart();
        pieChart.Append(new C.VaryColors { Val = true });

        foreach (var series in chart.Series)
        {
            var ser = new C.PieChartSeries();
            ser.Append(new C.Index { Val = (uint)series.Index });
            ser.Append(new C.Order { Val = (uint)series.Index });
            ser.Append(BuildSeriesName(series));

            if (chart.ChartType == ExcelChartType.PieExploded)
                ser.Append(new C.Explosion { Val = (uint)Math.Max(chart.ExplosionPercent, 25) });

            if (series.CategoriesFormula != null)
                ser.Append(BuildCategoryAxisData(series.CategoriesFormula));

            ser.Append(BuildNumericValues(series.ValuesFormula));
            pieChart.Append(ser);
        }

        pieChart.Append(BuildDataLabels(chart.ShowDataLabels));
        return pieChart;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Doughnut chart  (c:doughnutChart)
    // ═══════════════════════════════════════════════════════════════════════

    private static C.DoughnutChart BuildDoughnutChart(ExcelChart chart)
    {
        var doughnut = new C.DoughnutChart();
        doughnut.Append(new C.VaryColors { Val = true });

        foreach (var series in chart.Series)
        {
            var ser = new C.PieChartSeries(); // Doughnut reuses PieChartSeries
            ser.Append(new C.Index { Val = (uint)series.Index });
            ser.Append(new C.Order { Val = (uint)series.Index });
            ser.Append(BuildSeriesName(series));

            if (series.CategoriesFormula != null)
                ser.Append(BuildCategoryAxisData(series.CategoriesFormula));

            ser.Append(BuildNumericValues(series.ValuesFormula));
            doughnut.Append(ser);
        }

        doughnut.Append(BuildDataLabels(chart.ShowDataLabels));
        doughnut.Append(new C.HoleSize { Val = (byte)chart.DoughnutHoleSize });
        return doughnut;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Area chart  (c:areaChart)
    // ═══════════════════════════════════════════════════════════════════════

    private static C.AreaChart BuildAreaChart(ExcelChart chart)
    {
        var areaChart = new C.AreaChart();
        areaChart.Append(new C.Grouping { Val = GroupingToGroupingValues(chart.GetGroupingValue()) });

        foreach (var series in chart.Series)
        {
            var ser = new C.AreaChartSeries();
            ser.Append(new C.Index { Val = (uint)series.Index });
            ser.Append(new C.Order { Val = (uint)series.Index });
            ser.Append(BuildSeriesName(series));

            if (series.FillColorHex != null)
                ser.Append(BuildSolidFillShapeProperties(series.FillColorHex));

            if (series.CategoriesFormula != null)
                ser.Append(BuildCategoryAxisData(series.CategoriesFormula));

            ser.Append(BuildNumericValues(series.ValuesFormula));
            areaChart.Append(ser);
        }

        areaChart.Append(BuildDataLabels(chart.ShowDataLabels));
        areaChart.Append(new C.AxisId { Val = CatAxisId });
        areaChart.Append(new C.AxisId { Val = ValAxisId });

        return areaChart;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Scatter chart  (c:scatterChart)
    // ═══════════════════════════════════════════════════════════════════════

    private static C.ScatterChart BuildScatterChart(ExcelChart chart)
    {
        var scatterChart = new C.ScatterChart();

        scatterChart.Append(new C.ScatterStyle
        {
            Val = chart.ChartType switch
            {
                ExcelChartType.ScatterLines  => C.ScatterStyleValues.Line,
                ExcelChartType.ScatterSmooth => C.ScatterStyleValues.SmoothMarker,
                _                             => C.ScatterStyleValues.Marker
            }
        });

        scatterChart.Append(new C.VaryColors { Val = false });

        foreach (var series in chart.Series)
        {
            var ser = new C.ScatterChartSeries();
            ser.Append(new C.Index { Val = (uint)series.Index });
            ser.Append(new C.Order { Val = (uint)series.Index });
            ser.Append(BuildSeriesName(series));

            // Scatter uses XValues + YValues (not catAx / numRef)
            if (series.CategoriesFormula != null)
            {
                ser.Append(new C.XValues(
                    new C.NumberReference(
                        new C.Formula(series.CategoriesFormula))));
            }

            ser.Append(new C.YValues(
                new C.NumberReference(
                    new C.Formula(series.ValuesFormula))));

            scatterChart.Append(ser);
        }

        scatterChart.Append(BuildDataLabels(chart.ShowDataLabels));
        scatterChart.Append(new C.AxisId { Val = CatAxisId });
        scatterChart.Append(new C.AxisId { Val = ValAxisId });

        return scatterChart;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Radar chart  (c:radarChart)
    // ═══════════════════════════════════════════════════════════════════════

    private static C.RadarChart BuildRadarChart(ExcelChart chart)
    {
        var radarChart = new C.RadarChart();
        radarChart.Append(new C.RadarStyle
        {
            Val = chart.ChartType == ExcelChartType.RadarFilled
                ? C.RadarStyleValues.Filled
                : C.RadarStyleValues.Marker
        });
        radarChart.Append(new C.VaryColors { Val = false });

        foreach (var series in chart.Series)
        {
            var ser = new C.RadarChartSeries();
            ser.Append(new C.Index { Val = (uint)series.Index });
            ser.Append(new C.Order { Val = (uint)series.Index });
            ser.Append(BuildSeriesName(series));

            if (series.CategoriesFormula != null)
                ser.Append(BuildCategoryAxisData(series.CategoriesFormula));

            ser.Append(BuildNumericValues(series.ValuesFormula));
            radarChart.Append(ser);
        }

        radarChart.Append(BuildDataLabels(chart.ShowDataLabels));
        radarChart.Append(new C.AxisId { Val = CatAxisId });
        radarChart.Append(new C.AxisId { Val = ValAxisId });

        return radarChart;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Axes
    // ═══════════════════════════════════════════════════════════════════════

    private static C.CategoryAxis BuildCategoryAxis(ChartAxisOptions opts, ExcelChart chart)
    {
        var axis = new C.CategoryAxis();
        axis.Append(new C.AxisId { Val = CatAxisId });
        axis.Append(new C.Scaling(new C.Orientation
        {
            Val = C.OrientationValues.MinMax
        }));
        axis.Append(new C.Delete { Val = !opts.Visible });
        axis.Append(new C.AxisPosition { Val = C.AxisPositionValues.Bottom });

        if (opts.ShowMajorGridlines)
            axis.Append(new C.MajorGridlines());

        if (!string.IsNullOrEmpty(opts.Title))
            axis.Append(BuildAxisTitle(opts.Title!));

        // Tick marks
        axis.Append(new C.MajorTickMark { Val = TickMarkToOxml(opts.MajorTickMark) });
        axis.Append(new C.MinorTickMark { Val = TickMarkToOxml(opts.MinorTickMark) });

        // Tick label rotation
        if (opts.LabelRotation.HasValue || opts.FontSize.HasValue)
        {
            var txPr = BuildAxisTextProperties(opts.LabelRotation, opts.FontSize);
            axis.Append(txPr);
        }

        axis.Append(new C.NumberingFormat { FormatCode = "General", SourceLinked = true });
        axis.Append(new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
        axis.Append(new C.CrossingAxis { Val = ValAxisId });
        axis.Append(new C.Crosses { Val = C.CrossesValues.AutoZero });
        axis.Append(new C.AutoLabeled { Val = true });
        axis.Append(new C.LabelAlignment { Val = C.LabelAlignmentValues.Center });
        axis.Append(new C.LabelOffset { Val = 100 });

        return axis;
    }

    private static C.ValueAxis BuildValueAxis(ChartAxisOptions opts, ExcelChart chart)
    {
        var axis = new C.ValueAxis();
        axis.Append(new C.AxisId { Val = ValAxisId });

        // Scale
        var scaling = new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax });
        if (opts.MaxValue.HasValue) scaling.Append(new C.MaxAxisValue { Val = opts.MaxValue.Value });
        if (opts.MinValue.HasValue) scaling.Append(new C.MinAxisValue { Val = opts.MinValue.Value });
        axis.Append(scaling);

        axis.Append(new C.Delete { Val = !opts.Visible });
        axis.Append(new C.AxisPosition { Val = C.AxisPositionValues.Left });

        if (opts.ShowMajorGridlines)
            axis.Append(new C.MajorGridlines());
        if (opts.ShowMinorGridlines)
            axis.Append(new C.MinorGridlines());

        if (!string.IsNullOrEmpty(opts.Title))
            axis.Append(BuildAxisTitle(opts.Title!));

        // Tick marks
        axis.Append(new C.MajorTickMark { Val = TickMarkToOxml(opts.MajorTickMark) });
        axis.Append(new C.MinorTickMark { Val = TickMarkToOxml(opts.MinorTickMark) });

        // Number format
        string numFmt = opts.NumberFormat ?? "General";
        axis.Append(new C.NumberingFormat { FormatCode = numFmt, SourceLinked = string.IsNullOrEmpty(opts.NumberFormat) });

        axis.Append(new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
        axis.Append(new C.CrossingAxis { Val = CatAxisId });
        axis.Append(new C.Crosses { Val = C.CrossesValues.AutoZero });
        axis.Append(new C.CrossBetween { Val = C.CrossBetweenValues.Between });

        if (opts.MajorUnit.HasValue)
            axis.Append(new C.MajorUnit { Val = opts.MajorUnit.Value });

        if (opts.LabelRotation.HasValue || opts.FontSize.HasValue)
            axis.Append(BuildAxisTextProperties(opts.LabelRotation, opts.FontSize));

        return axis;
    }

    private static C.Title BuildAxisTitle(string text)
    {
        return new C.Title(
            new C.ChartText(
                new C.RichText(
                    new A.BodyProperties { Rotation = 0 },
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.Run(new A.Text(text))))),
            new C.Overlay { Val = false });
    }

    private static C.TextProperties BuildAxisTextProperties(int? rotation, double? fontSize)
    {
        var bodyProps = new A.BodyProperties();
        if (rotation.HasValue)
        {
            // OOXML rotation: 60000 units per degree
            bodyProps.Rotation = rotation.Value * 60000;
        }

        var defRPr = new A.DefaultRunProperties();
        if (fontSize.HasValue)
            defRPr.FontSize = (int)(fontSize.Value * 100); // hundredths of a point

        return new C.TextProperties(
            bodyProps,
            new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(defRPr)));
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Legend
    // ═══════════════════════════════════════════════════════════════════════

    private static C.Legend BuildLegend(ChartLegendOptions opts)
    {
        var legend = new C.Legend();

        legend.Append(new C.LegendPosition
        {
            Val = opts.Position switch
            {
                ChartLegendPosition.Top      => C.LegendPositionValues.Top,
                ChartLegendPosition.Left     => C.LegendPositionValues.Left,
                ChartLegendPosition.Right    => C.LegendPositionValues.Right,
                ChartLegendPosition.TopRight => C.LegendPositionValues.TopRight,
                _                             => C.LegendPositionValues.Bottom
            }
        });

        legend.Append(new C.Overlay { Val = opts.OverlayChart });

        if (opts.FontSize.HasValue)
        {
            legend.Append(new C.TextProperties(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(
                    new A.ParagraphProperties(
                        new A.DefaultRunProperties
                        {
                            FontSize = (int)(opts.FontSize.Value * 100)
                        }))));
        }

        return legend;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Shared series helpers
    // ═══════════════════════════════════════════════════════════════════════

    private static C.SeriesText BuildSeriesName(ChartSeries series)
    {
        if (series.NameIsFormula)
        {
            // Formula reference: Sheet1!$A$1
            return new C.SeriesText(
                new C.StringReference(
                    new C.Formula(series.Name)));
        }

        // Literal string name - use StringLiteral with proper structure
        // Note: C.NumericValue is for numbers, not text. Excel tolerates it but
        // proper OOXML uses StringLiteral for text values.
        return new C.SeriesText(
            new C.StringLiteral(
                new C.PointCount { Val = 1 },
                new C.StringPoint(
                    new C.NumericValue(series.Name)) { Index = 0 }));
    }

    private static C.CategoryAxisData BuildCategoryAxisData(string formula)
    {
        return new C.CategoryAxisData(
            new C.StringReference(
                new C.Formula(formula)));
    }

    private static C.Values BuildNumericValues(string formula)
    {
        return new C.Values(
            new C.NumberReference(
                new C.Formula(formula)));
    }

    /// <summary>
    /// Builds a solid fill shape properties element for series color override.
    /// </summary>
    /// <param name="argbHex">
    /// ARGB hex color (8 chars) or RGB hex color (6 chars).
    /// Note: RgbColorModelHex in OOXML only supports RGB (6 chars).
    /// If ARGB is provided, the alpha channel is stripped because
    /// chart series fills do not support transparency in standard OOXML.
    /// </param>
    private static C.ChartShapeProperties BuildSolidFillShapeProperties(string argbHex)
    {
        // RgbColorModelHex requires 6-char RGB, not 8-char ARGB
        // Strip alpha prefix if present (first 2 chars of 8-char ARGB)
        string rgbHex = argbHex.Length == 8 ? argbHex[2..] : argbHex;
        
        return new C.ChartShapeProperties(
            new A.SolidFill(
                new A.RgbColorModelHex { Val = rgbHex }));
    }

    private static C.DataLabels BuildDataLabels(bool show)
    {
        if (!show)
        {
            return new C.DataLabels(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false });
        }

        return new C.DataLabels(
            new C.ShowLegendKey { Val = false },
            new C.ShowValue { Val = true },
            new C.ShowCategoryName { Val = false },
            new C.ShowSeriesName { Val = false },
            new C.ShowPercent { Val = false },
            new C.ShowBubbleSize { Val = false });
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Enum conversions
    // ═══════════════════════════════════════════════════════════════════════

    private static C.BarGroupingValues GroupingToBarGrouping(string grouping)
        => grouping switch
        {
            "stacked"        => C.BarGroupingValues.Stacked,
            "percentStacked" => C.BarGroupingValues.PercentStacked,
            _                 => C.BarGroupingValues.Clustered
        };

    private static C.GroupingValues GroupingToGroupingValues(string grouping)
        => grouping switch
        {
            "stacked"        => C.GroupingValues.Stacked,
            "percentStacked" => C.GroupingValues.PercentStacked,
            "standard"       => C.GroupingValues.Standard,
            _                 => C.GroupingValues.Standard
        };

    private static C.MarkerStyleValues MarkerStyleToOxml(ChartMarkerStyle style)
        => style switch
        {
            ChartMarkerStyle.None     => C.MarkerStyleValues.None,
            ChartMarkerStyle.Circle   => C.MarkerStyleValues.Circle,
            ChartMarkerStyle.Square   => C.MarkerStyleValues.Square,
            ChartMarkerStyle.Diamond  => C.MarkerStyleValues.Diamond,
            ChartMarkerStyle.Triangle => C.MarkerStyleValues.Triangle,
            ChartMarkerStyle.Star     => C.MarkerStyleValues.Star,
            ChartMarkerStyle.Plus     => C.MarkerStyleValues.Plus,
            ChartMarkerStyle.X        => C.MarkerStyleValues.X,
            _                          => C.MarkerStyleValues.Auto
        };

    private static C.TickMarkValues TickMarkToOxml(ChartTickMark tickMark)
        => tickMark switch
        {
            ChartTickMark.None    => C.TickMarkValues.None,
            ChartTickMark.Inside  => C.TickMarkValues.Inside,
            ChartTickMark.Outside => C.TickMarkValues.Outside,
            ChartTickMark.Cross   => C.TickMarkValues.Cross,
            _                      => C.TickMarkValues.Outside
        };
}
