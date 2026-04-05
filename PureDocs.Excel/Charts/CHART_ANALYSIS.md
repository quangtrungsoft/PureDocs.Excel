# Phân Tích Kỹ Thuật — TVE.PureDocs.Excel Chart Feature

## 1. Phân Tích Kiến Trúc Codebase Hiện Tại

### 1.1 Object Model Overview

```
Workbook  (sealed, IDisposable)
  ├── SharedStringManager   [internal]
  ├── StyleManager          [internal]
  └── WorksheetCollection
        └── Worksheet[]  (sealed)
              ├── Cell (sealed)
              │     ├── StyleManager ref
              │     └── SharedStringManager ref
              └── CellStyle (sealed)
                    ├── ExcelFont
                    ├── ExcelFill
                    ├── ExcelBorder
                    ├── ExcelAlignment
                    └── ExcelNumberFormat
```

### 1.2 Design Patterns Được Áp Dụng

| Pattern | Class | Mô tả |
|---|---|---|
| **Manager / Service** | `StyleManager`, `SharedStringManager` | Tách biệt logic OpenXML khỏi domain model |
| **Wrapper** | `Cell`, `Worksheet`, `Workbook` | Bọc OpenXML SDK bằng clean domain API |
| **Fluent Builder** | `Cell`, `CellStyle` | Method chaining trả về `this` |
| **Deduplication Cache** | `StyleManager` | XML-key → index cache cho font/fill/border |
| **Factory Method** | `Workbook.Create()`, `Workbook.Open()` | Kiểm soát construction |
| **Two-Phase Construction** | `InitializeWorkbook` + constructor | OpenXML init tách biệt khỏi object init |

### 1.3 Điểm Mạnh Kỹ Thuật

**SharedStringManager — O(1) Lookup:**
```csharp
// Dictionary<string,int> + List<string>: bidirectional O(1)
private readonly Dictionary<string, int> _stringToIndex;
private readonly List<string> _indexToString;

// AddOrGetString: O(1) average — no linear scan
if (_stringToIndex.TryGetValue(value, out int existingIndex))
    return existingIndex;
```

**StyleManager — XML-key deduplication:**
```csharp
// Mỗi component (font, fill, border) serialize thành XML string → làm cache key
string key = font.OuterXml;
if (_fontCache.TryGetValue(key, out uint existingIndex))
    return existingIndex;
// Đảm bảo không tạo duplicate style entries
```

**Ordered cell insertion — tránh invalid OOXML:**
```csharp
// FindOrCreateCell: duy trì thứ tự cột tăng dần (required by OOXML schema)
if (string.Compare(existingCell.CellReference?.Value, cellReference) > 0)
    row.InsertBefore(cell, refCell);
```

**MergeCells positioning:**
```csharp
// MergeCells phải đặt AFTER SheetData — code xử lý đúng ordering
worksheet.InsertAfter(mergeCells, sheetData);
```

### 1.4 Điểm Cần Cải Thiện (Technical Debt)

| Vấn đề | Vị trí | Tác động |
|---|---|---|
| `WorksheetCollection.RemoveAt()` không xóa OpenXML part | `WorksheetCollection.cs:54` | Memory leak, file vẫn chứa deleted sheet |
| `GetStatistics()` hardcode hit rate = 100 | `SharedStringManager.cs:84` | Monitoring không có giá trị |
| `IsDateFormatIndex()` check regex trên format code có thể false positive | `Cell.cs:131` | Cell có chứ "m" → sai nhận là date |
| `SaveStylesheet()` gọi mỗi operation | `StyleManager.cs` | I/O overhead khi batch styling |
| `_workbook` nullable trên `Worksheet` | `Worksheet.cs:16` | Chart extension cần `WorksheetPart` — cần thêm accessor |

---

## 2. Kiến Trúc Chart Feature

### 2.1 OOXML Chart Structure

Hiểu cấu trúc OOXML là nền tảng để implement đúng:

```
WorksheetPart
  └── DrawingsPart                    ← 1 per worksheet
        ├── WorksheetDrawing (xdr:wsDr)
        │     └── TwoCellAnchor       ← 1 per chart (position on grid)
        │           ├── From (col, row, offsets)
        │           ├── To   (col, row, offsets)
        │           ├── GraphicFrame
        │           │     └── c:chart r:id="rId1"  ← reference to ChartPart
        │           └── ClientData
        └── ChartPart (rId1)           ← 1 per chart
              └── ChartSpace (c:chartSpace)
                    └── Chart (c:chart)
                          ├── Title
                          ├── PlotArea
                          │     ├── Layout
                          │     ├── c:barChart | c:lineChart | c:pieChart | ...
                          │     │     ├── BarDirection / Grouping
                          │     │     ├── c:ser (series) ×N
                          │     │     │     ├── Index, Order
                          │     │     │     ├── SeriesText (name)
                          │     │     │     ├── CategoryAxisData (c:cat)
                          │     │     │     └── Values (c:val → NumberReference → Formula)
                          │     │     ├── DataLabels
                          │     │     ├── AxisId (catAx ref)
                          │     │     └── AxisId (valAx ref)
                          │     ├── CategoryAxis (c:catAx)
                          │     └── ValueAxis (c:valAx)
                          └── Legend
```

### 2.2 Class Diagram Chart Feature

```
ExcelChart                     (public sealed — fluent builder)
  ├── ChartType: ExcelChartType
  ├── Position: ChartPosition
  ├── Series: List<ChartSeries>
  ├── CategoryAxis: ChartAxisOptions
  ├── ValueAxis: ChartAxisOptions
  └── Legend: ChartLegendOptions

ChartSeries                    (public sealed — fluent)
  ├── Name: string
  ├── ValuesFormula: string
  ├── CategoriesFormula: string?
  ├── FillColorHex: string?
  └── MarkerStyle: ChartMarkerStyle

ChartPosition                  (public sealed — value object)
  ├── From(row,col,toRow,toCol) — 1-based factory
  └── FromZeroBased(...)        — 0-based raw OOXML

ChartAxisOptions               (public sealed — configurable)
  ├── Title, Visible
  ├── ShowMajorGridlines, ShowMinorGridlines
  ├── NumberFormat, MinValue, MaxValue, MajorUnit
  └── LabelRotation, FontSize

ChartLegendOptions             (public sealed — configurable)
  ├── Visible, Position
  ├── OverlayChart, FontSize
  └── ...

ChartWriter                    (internal sealed — OpenXML engine)
  ├── WriteChart(WorksheetPart, ExcelChart) → rId
  ├── EnsureDrawingsPart()
  ├── BuildTwoCellAnchor()
  ├── BuildChartSpace()
  ├── BuildPlotArea() [dispatch table]
  ├── BuildBarColumnChart / BuildLineChart / BuildPieChart / ...
  └── BuildCategoryAxis / BuildValueAxis / BuildLegend

WorksheetChartExtensions       (public static — extension methods)
  ├── AddChart(worksheet, type, position) → ExcelChart
  ├── CommitChart(worksheet, chart) → ExcelChart
  └── AddAndCommitChart(worksheet, type, pos, configure)
```

### 2.3 Luồng Xử Lý (Sequence Diagram)

```
User Code                   Worksheet Extension    ChartWriter          OpenXML SDK
   │                               │                    │                    │
   │── AddChart(type, pos) ───────►│                    │                    │
   │◄─ ExcelChart instance ────────│                    │                    │
   │                               │                    │                    │
   │── chart.SetTitle("...") ─────►│ (ExcelChart fluent)│                    │
   │── chart.AddSeries("Rev","...") ►                   │                    │
   │        .SetCategories("...") ─►                   │                    │
   │── chart.ConfigureValueAxis() ─►                   │                    │
   │                               │                    │                    │
   │── CommitChart(chart) ─────────►│                   │                    │
   │                               │── WriteChart() ───►│                    │
   │                               │                    │── EnsureDrawingsPart►│
   │                               │                    │◄─ DrawingsPart ────│
   │                               │                    │                    │
   │                               │                    │── AddNewPart<ChartPart>►│
   │                               │                    │◄─ chartPart ───────│
   │                               │                    │                    │
   │                               │                    │── BuildChartSpace()│
   │                               │                    │   [dispatch table] │
   │                               │                    │── chartPart.Save() ►│
   │                               │                    │                    │
   │                               │                    │── BuildTwoCellAnchor│
   │                               │                    │── drawing.Append() ►│
   │◄─ ExcelChart ─────────────────│◄──────────────────│                    │
```

---

## 3. Thiết Kế Chi Tiết Các File

### 3.1 `ExcelChartType.cs`

Enum với 25+ chart types, grouped theo category. Tránh dùng flags (không cần combination).

Key design decisions:
- Mỗi type name embed cả grouping (e.g., `ColumnClustered`, `ColumnStacked`) — tránh cần thêm grouping property riêng cho common cases
- `ExcelChart.GetGroupingValue()` extract grouping từ type name — DRY
- `IsPieOrDoughnut()`, `IsBarChart()`, `IsScatter()` helper properties để ChartWriter dispatch logic clean

### 3.2 `ChartSeries.cs`

```csharp
// Key design: formula string làm reference, không phải IRange
// Tại sao? Vì chart data range tồn tại độc lập với cell objects
// và được lưu dưới dạng formula string trong OOXML:
//   <c:f>Sheet1!$B$2:$B$13</c:f>

// Formula normalization: strip leading '=' để consistent
private static string NormalizeFormula(string formula)
    => formula.TrimStart('=');

// NameIsFormula flag: phân biệt literal vs cell reference
// "Revenue" → <c:v>Revenue</c:v>   (NumericValue với text)
// "Sheet1!$B$1" → <c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef>
```

### 3.3 `ChartOptions.cs`

**ChartPosition — Coordinate System Mapping:**
```
Excel UI (1-based)     OOXML xdr:twoCellAnchor (0-based)
    Row 1, Col A    →      rowId=0, colId=0
    Row 2, Col B    →      rowId=1, colId=1
    
// Factory method handles the conversion:
FromRow    = row - 1,    // 1-based → 0-based
FromColumn = col - 1
```

**ChartAxisOptions — Scale configuration:**
```csharp
// MinValue/MaxValue/MajorUnit chỉ áp dụng cho ValueAxis
// Khi null → Excel auto-scale (recommended default)
// OOXML: <c:scaling><c:max val="100"/><c:min val="0"/></c:scaling>
```

### 3.4 `ExcelChart.cs`

**Fluent Pattern Implementation:**
```csharp
// Mọi method đều return ExcelChart để support chaining
// EXCEPT AddSeries — trả về ChartSeries để config series cụ thể

var chart = ws.AddChart(ExcelChartType.ColumnClustered, pos)
    .SetTitle("Sales")
    .AddSeries("Q1", "Sheet1!$B$2:$B$5")      // ← ChartSeries
        .SetCategories("Sheet1!$A$2:$A$5")
        .SetFillColor("#4472C4")
    // Tiếp tục config chart (không phải series):
    // Cần gọi lại ws.CommitChart(chart) vì fluent chain đã break
```

> **Note:** Hạn chế của chain-breaking: sau `AddSeries()` trả về `ChartSeries`, user không thể tiếp tục chain `ExcelChart`. Giải pháp là dùng `ConfigureCategoryAxis(a => ...)` pattern thay vì returning different types.

**GapWidth và Overlap:**
```
GapWidth:  0 = bars touch, 150 = Excel default, 500 = max gap
Overlap: -100 = max gap within cluster, 0 = no overlap, 100 = full overlap

// OOXML: <c:gapWidth val="150"/>  <c:overlap val="0"/>
```

### 3.5 `ChartWriter.cs` — Core Engine

**DrawingsPart initialization — critical OOXML ordering:**
```csharp
// Worksheet element children MUST follow this order (OOXML schema):
// SheetData → ConditionalFormatting → DataValidations →
// Hyperlinks → PrintOptions → PageMargins → PageSetup →
// Drawing → ...
//
// Implementation: Append <drawing> at end — safe because it's
// always the last structural element before optional metadata
wsElement.Append(drawingRef);
```

**Axis ID constants:**
```csharp
// Axis IDs liên kết chart element với axis element trong OOXML
// Phải consistent giữa c:barChart/c:lineChart và c:catAx/c:valAx
private const uint CatAxisId = 2093507336;
private const uint ValAxisId = 2093510856;

// c:barChart chứa:   <c:axId val="2093507336"/> <c:axId val="2093510856"/>
// c:catAx chứa:      <c:axId val="2093507336"/>
//                    <c:crossAx val="2093510856"/>
// c:valAx chứa:      <c:axId val="2093510856"/>
//                    <c:crossAx val="2093507336"/>
```

**Series name encoding:**
```csharp
// Literal name → c:ser/c:tx/c:v
BuildSeriesName("Revenue") →
  <c:tx><c:v>Revenue</c:v></c:tx>

// Formula reference → c:ser/c:tx/c:strRef/c:f
BuildSeriesName("Sheet1!$B$1", isFormula:true) →
  <c:tx><c:strRef><c:f>Sheet1!$B$1</c:f></c:strRef></c:tx>
```

**Rotation encoding:**
```csharp
// OOXML TextBody rotation: 60000 units per degree
// 45° → 2,700,000 units
// -45° → -2,700,000 units
bodyProps.Rotation = rotationDegrees * 60000;
```

**Font size encoding:**
```csharp
// OOXML font size: hundredths of a point
// 12pt → 1200
defRPr.FontSize = (int)(fontSize * 100);
```

### 3.6 `WorksheetChartExtensions.cs`

**ConditionalWeakTable pattern:**
```csharp
// ChartWriter là stateful (track _nextShapeId per worksheet)
// Dùng ConditionalWeakTable thay vì field trên Worksheet để:
//   1. Không modify sealed class
//   2. GC-friendly: khi Worksheet bị GC thì ChartWriter cũng được collect
//   3. Thread-safe cho concurrent worksheet operations

private static readonly ConditionalWeakTable<Worksheet, ChartWriter> _writerTable = new();
var writer = _writerTable.GetOrCreateValue(worksheet);
```

**Validation guard:**
```csharp
// CommitChart throw trước khi write nếu không có series
if (chart.Series.Count == 0)
    throw new InvalidOperationException("Chart must have at least one series.");
// Tại sao? Một chart không có series → invalid OOXML, Excel sẽ repair/corrupt file
```

---

## 4. Modification cần thiết trên Worksheet.cs

Thêm một method internal vào class `Worksheet` (không phá vỡ public API):

```csharp
// Thêm vào Worksheet.cs, phần "Internal helpers"

/// <summary>
/// Returns the underlying <see cref="WorksheetPart"/> for chart embedding.
/// Called exclusively by <see cref="WorksheetChartExtensions"/>.
/// </summary>
internal WorksheetPart GetWorksheetPart() => _worksheetPart;
```

---

## 5. Ví Dụ Sử Dụng

### 5.1 Column Chart cơ bản

```csharp
using TVE.PureDocs.Excel;
using TVE.PureDocs.Excel.Charts;

using var wb = Workbook.Create();
var ws = wb.AddWorksheet("Sales");

// Populate data
ws.GetCell("A1").SetValue("Month");
ws.GetCell("B1").SetValue("Revenue");
ws.GetCell("C1").SetValue("Cost");

string[] months = { "Jan","Feb","Mar","Apr","May","Jun" };
double[] rev =  { 120000, 135000, 118000, 152000, 143000, 160000 };
double[] cost = { 85000,  91000,  80000,  102000, 96000,  108000 };

for (int i = 0; i < 6; i++)
{
    ws.GetCell(i + 2, 1).SetValue(months[i]);
    ws.GetCell(i + 2, 2).SetValue(rev[i]);
    ws.GetCell(i + 2, 3).SetValue(cost[i]);
}

// Create chart
var chart = ws.AddChart(
    ExcelChartType.ColumnClustered,
    ChartPosition.From(row: 2, col: 5, toRow: 18, toCol: 13));

chart
    .SetTitle("Monthly Sales Performance")
    .SetGapWidth(80)
    .ConfigureValueAxis(a => a
        .SetNumberFormat("#,##0")
        .ShowGridlines()
        .SetTitle("Amount (VND)"))
    .ConfigureCategoryAxis(a => a
        .SetTitle("Month"))
    .ConfigureLegend(l => l.SetPosition(ChartLegendPosition.Bottom));

// Series 1
chart.AddSeries("Revenue", "Sales!$B$2:$B$7")
    .SetCategories("Sales!$A$2:$A$7")
    .SetFillColor("#4472C4");

// Series 2
chart.AddSeries("Cost", "Sales!$C$2:$C$7")
    .SetCategories("Sales!$A$2:$A$7")
    .SetFillColor("#ED7D31");

ws.CommitChart(chart);
wb.SaveAs("sales_report.xlsx");
```

### 5.2 Line Chart với markers

```csharp
var lineChart = ws.AddChart(
    ExcelChartType.LineWithMarkers,
    ChartPosition.From(2, 2, 16, 10));

lineChart.SetTitle("Trend Analysis");

lineChart.AddSeries("Sheet1!$B$1", "Sheet1!$B$2:$B$13", nameIsFormula: true)
    .SetCategories("Sheet1!$A$2:$A$13")
    .SetMarkerStyle(ChartMarkerStyle.Circle);

lineChart.ConfigureValueAxis(a => a
    .SetScale(min: 0, max: 200000, majorUnit: 50000)
    .SetNumberFormat("#,##0"));

ws.CommitChart(lineChart);
```

### 5.3 Pie Chart

```csharp
ws.AddAndCommitChart(
    ExcelChartType.Pie,
    ChartPosition.From(2, 6, 18, 12),
    chart =>
    {
        chart.SetTitle("Market Share")
             .ShowLabels()
             .ConfigureLegend(l => l.SetPosition(ChartLegendPosition.Right));

        chart.AddSeries("Market Share", "Sheet1!$B$2:$B$6")
            .SetCategories("Sheet1!$A$2:$A$6");
    });
```

### 5.4 Doughnut Chart

```csharp
var doughnut = ws.AddChart(ExcelChartType.Doughnut,
    ChartPosition.From(2, 2, 16, 9));

doughnut
    .SetTitle("Portfolio Breakdown")
    .SetDoughnutHoleSize(60)
    .ShowLabels()
    .ConfigureLegend(l => l.SetPosition(ChartLegendPosition.Right));

doughnut.AddSeries("Portfolio", "Sheet1!$B$2:$B$5")
    .SetCategories("Sheet1!$A$2:$A$5");

ws.CommitChart(doughnut);
```

### 5.5 Multiple charts trên cùng worksheet

```csharp
// Chart 1 — top section
var chart1 = ws.AddChart(ExcelChartType.ColumnClustered,
    ChartPosition.From(1, 6, 16, 13));
chart1.AddSeries("Revenue", "Data!$B$2:$B$13")
    .SetCategories("Data!$A$2:$A$13");
ws.CommitChart(chart1);

// Chart 2 — bottom section
var chart2 = ws.AddChart(ExcelChartType.Line,
    ChartPosition.From(17, 6, 32, 13));
chart2.AddSeries("Growth %", "Data!$C$2:$C$13")
    .SetCategories("Data!$A$2:$A$13");
chart2.ConfigureValueAxis(a => a.SetNumberFormat("0.0%"));
ws.CommitChart(chart2);
```

---

## 6. Mapping Chart Types → OOXML Elements

| ExcelChartType | OOXML Element | c:barDir / c:style |
|---|---|---|
| ColumnClustered | `c:barChart` | col / clustered |
| ColumnStacked | `c:barChart` | col / stacked |
| ColumnStackedFull | `c:barChart` | col / percentStacked |
| BarClustered | `c:barChart` | bar / clustered |
| BarStacked | `c:barChart` | bar / stacked |
| Line | `c:lineChart` | grouping: standard |
| LineWithMarkers | `c:lineChart` | + marker symbol |
| LineSmooth | `c:lineChart` | smooth=true |
| Pie | `c:pieChart` | varyColors=true |
| PieExploded | `c:pieChart` | + c:explosion |
| Doughnut | `c:doughnutChart` | holeSize |
| Area | `c:areaChart` | standard |
| AreaStacked | `c:areaChart` | stacked |
| ScatterMarkers | `c:scatterChart` | marker style |
| ScatterLines | `c:scatterChart` | line style |
| Radar | `c:radarChart` | marker style |
| RadarFilled | `c:radarChart` | filled style |

---

## 7. File Structure

```
TVE.PureDocs.Excel/
├── (existing files)
│
└── Charts/
    ├── ExcelChartType.cs           ← Enums: ExcelChartType, ChartGrouping, etc.
    ├── ChartSeries.cs              ← Series data binding + fluent config
    ├── ChartOptions.cs             ← ChartAxisOptions, ChartLegendOptions, ChartPosition
    ├── ExcelChart.cs               ← Main public chart builder
    ├── ChartWriter.cs              ← Internal OpenXML serializer
    ├── WorksheetChartExtensions.cs ← Extension methods on Worksheet
    └── WorksheetInternals.cs       ← Documents the internal accessor requirement
```

**Required modification to existing file:**
```
Worksheet.cs → add:  internal WorksheetPart GetWorksheetPart() => _worksheetPart;
```

---

## 8. NuGet Dependencies

```xml
<!-- Chart feature requires these packages (đã có trong DocumentFormat.OpenXml) -->
<PackageReference Include="DocumentFormat.OpenXml" Version="3.x" />

<!-- Chart namespaces sử dụng: -->
<!-- DocumentFormat.OpenXml.Drawing.Charts       (c:) -->
<!-- DocumentFormat.OpenXml.Drawing.Spreadsheet  (xdr:) -->
<!-- DocumentFormat.OpenXml.Drawing              (a:) -->
<!-- DocumentFormat.OpenXml.Packaging            (ChartPart, DrawingsPart) -->
```

Không cần thêm NuGet package — chart support nằm trong `DocumentFormat.OpenXml` core.
