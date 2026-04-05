# TVE.PureDocs.Excel - Thư viện Excel .NET Hiệu năng cao

**TVE.PureDocs.Excel** là một thư viện .NET mạnh mẽ, thuần túy (pure .NET) được thiết kế để tạo, đọc, ghi và tính toán các tệp Excel (.xlsx) với hiệu suất cực cao. Thư viện này hoạt động hoàn toàn trong bộ nhớ, không yêu cầu cài đặt Microsoft Office hay Excel Interop, tương thích hoàn hảo với các môi trường Cloud, Linux, Docker và Windows.

---

## 🌟 Điểm nổi bật (Key Features)

*   🚀 **Hiệu suất vượt trội**: Thao tác hàng triệu ô tính chỉ trong vài giây với cơ chế quản lý bộ nhớ thông minh.
*   🧮 **Formula Engine mạnh mẽ**: Hỗ trợ hơn 100+ hàm Excel thông dụng với khả năng **Smart Recalculation** (chỉ tính lại những ô bị ảnh hưởng).
*   🎨 **Fluent Styling API**: Thiết kế định dạng ô tính (Font, Màu sắc, Viền, Căn lề) một cách trực quan và gọn gàng.
*   📊 **Hỗ trợ Biểu đồ**: Tạo nhanh 22 loại biểu đồ (Column, Bar, Line, Pie, Radar...) với API cấu hình chuyên sâu.
*   🏗️ **Bố cục linh hoạt**: Merge cells, Freeze panes, Auto-fit cột, AutoFilter, Dropdown Validation.
*   🧩 **Deduplication tự động**: Tối ưu dung lượng file bằng cách tự động loại bỏ các định dạng trùng lặp và sử dụng Shared String Table.

---

## ⚙️ Cài đặt

Thư viện hỗ trợ .NET 7, .NET 8 và .NET 9.

```bash
dotnet add package TVE.PureDocs.Excel
```
*(Hoặc cài đặt qua NuGet Package Manager)*

---

## 🚀 Hướng dẫn sử dụng nhanh

### 1. Tạo Workbook và Ghi dữ liệu

```csharp
using TVE.PureDocs.Excel;

// 1. Tạo Workbook mới
using var workbook = Workbook.Create();

// 2. Thêm Worksheet
var sheet = workbook.AddWorksheet("Báo cáo doanh thu");

// 3. Ghi dữ liệu theo Indexer (A1, B1...)
sheet["A1"].SetValue("Sản phẩm");
sheet["B1"].SetValue("Số lượng");
sheet["C1"].SetValue("Đơn giá");
sheet["D1"].SetValue("Tổng cộng");

// Ghi dữ liệu theo Row/Column (1-based)
sheet[2, 1].SetValue("Laptop Dell XPS");
sheet[2, 2].SetValue(5);
sheet[2, 3].SetValue(1200.50);

// 4. Lưu file
workbook.SaveAs("BaoCao.xlsx");
```

### 2. Định dạng nâng cao (Fluent API)

```csharp
// Định dạng Header chuyên nghiệp
sheet.GetRange("A1:D1").Cells.ForEach(cell => 
{
    cell.Style
        .SetBold(true)
        .SetFontSize(12)
        .SetFontColor(ExcelColor.White)
        .SetBackgroundColor(ExcelColor.DarkBlue)
        .SetHorizontalAlignment(ExcelHorizontalAlignment.Center)
        .SetAllBorders(ExcelBorderStyle.Medium);
});

// Định dạng số tiền
sheet["C2"].SetNumberFormat("#,##0.00 $");
```

---

## 📖 Tính năng chi tiết

### 🧮 Engine Tính toán Công thức

Thư viện không chỉ lưu công thức mà còn có thể **tính toán giá trị ngay lập tức** mà không cần mở Excel.

```csharp
// Gán công thức
sheet["D2"].SetFormula("=B2*C2");

// Tính toán thông minh (Smart Recalculate)
// Chỉ những ô phụ thuộc vào B2 hoặc C2 mới được tính lại
sheet.SmartRecalculate();

// Lấy kết quả sau tính toán
var total = sheet["D2"].GetValue(); 
```

### 📊 Thêm Biểu đồ (Charts)

Hỗ trợ 22 loại biểu đồ với khả năng cấu hình chi tiết (Tiêu đề, Trục, Chú giải).

```csharp
using TVE.PureDocs.Excel.Charts;

// Tạo biểu đồ cột
var chart = sheet.AddChart(ExcelChartType.ColumnClustered, 
                          ChartPosition.From(row: 5, col: 1, toRow: 20, toCol: 8))
    .SetTitle("Biểu đồ doanh thu năm 2026")
    .AddSeries("Doanh thu", "Báo cáo doanh thu!$B$2:$B$10")
        .SetCategories("Báo cáo doanh thu!$A$2:$A$10")
    .ConfigureValueAxis(axis => axis.SetTitle("USD").ShowGridlines())
    .ConfigureLegend(legend => legend.SetPosition(ChartLegendPosition.Bottom));

// Commit để viết dữ liệu vào file
sheet.CommitChart(chart);
```

### 📑 Thao tác Vùng dữ liệu (Range)

```csharp
var range = sheet.GetRange("A1:D100");

// Ghi dữ liệu hàng loạt từ mảng 2 chiều
object[,] data = { { "A", 1 }, { "B", 2 } };
range.SetValues(data);

// Tự động giãn cột theo nội dung
range.AutoFit();
```

---

## 🛠️ Tính năng Nâng cao (Advanced Features)

Dưới đây là hướng dẫn chi tiết cho các tính năng chuyên sâu của thư viện.

### 1. Ràng buộc dữ liệu (Data Validation - Dropdown)

Bạn có thể tạo các danh sách thả xuống (Dropdown) để giới hạn giá trị nhập vào một ô hoặc vùng ô.

```csharp
// Cách 1: Dropdown từ danh sách cố định (nguyên văn)
// Lưu ý: Chuỗi list phải nằm trong dấu ngoặc kép "\" value1,value2 \""
sheet.AddDropdownValidation("B2:B10", "\"Nam,Nữ,Khác\"");

// Cách 2: Dropdown tham chiếu từ một dải ô khác (ví dụ từ sheet 'Danh mục')
sheet.AddDropdownValidation("C2:C100", "'Danh mục'!$A$2:$A$50");
```

### 2. Biểu đồ đa dạng & Cấu hình chi tiết (Advanced Charts)

Thư viện hỗ trợ 22 loại biểu đồ. Bạn có thể thêm nhiều Series (chuỗi dữ liệu) và cấu hình trục tọa độ.

```csharp
using TVE.PureDocs.Excel.Charts;

// Tạo biểu đồ đường (Line Chart) với 2 chuỗi dữ liệu
var chart = sheet.AddChart(ExcelChartType.LineWithMarkers, ChartPosition.From(5, 5, 25, 15))
    .SetTitle("So sánh Doanh thu & Chi phí")
    .Set3D(false)
    
    // Thêm Series 1: Doanh thu
    .AddSeries("Doanh thu", "Sheet1!$B$2:$B$13")
        .SetCategories("Sheet1!$A$2:$A$13")
        .SetFillColor("#4472C4") // Màu xanh
    
    // Thêm Series 2: Chi phí
    .AddSeries("Chi phí", "Sheet1!$C$2:$C$13")
        .SetFillColor("#ED7D31"); // Màu cam

// Cấu hình trục giá trị (Trục Y)
chart.ConfigureValueAxis(axis => {
    axis.SetTitle("VNĐ")
        .ShowGridlines(true)
        .SetMajorTickMark(TickMarkValues.Outside);
});

// Cấu hình Chú thích (Legend)
chart.ConfigureLegend(legend => {
    legend.SetPosition(ChartLegendPosition.Right)
          .SetOverlay(false);
});

// Lưu cấu hình vào file
sheet.CommitChart(chart);
```

### 3. Thao tác dữ liệu khối lượng lớn (Bulk Range Operations)

Để đạt hiệu suất cao nhất khi ghi hàng ngàn dòng dữ liệu, hãy sử dụng mảng 2 chiều và `SetValues`.

```csharp
// Chuẩn bị dữ liệu mảng 2 chiều
object[,] data = new object[100, 4];
for (int i = 0; i < 100; i++) {
    data[i, 0] = $"SP-{i}";     // Text
    data[i, 1] = i + 1;         // Number
    data[i, 2] = 15.5 * i;      // Decimal
    data[i, 3] = DateTime.Now;  // DateTime
}

// Ghi toàn bộ khối dữ liệu vào sheet chỉ với 1 lệnh
sheet.GetRange("A2:D101").SetValues(data);

// Tự động căn chỉnh độ rộng cho tất cả cột trong vùng
sheet.GetRange("A1:D101").AutoFit();
```

### 4. Công thức nâng cao & Tham chiếu (Formula Engine)

Hỗ trợ tham chiếu chéo giữa các Sheet và đặt tên vùng ô (Named Ranges).

```csharp
using TVE.PureDocs.Excel.Formulas;

// 1. Tham chiếu chéo Sheet
sheet["A1"].SetFormula("='Tháng 1'!B2 + 'Tháng 2'!B2");

// 2. Sử dụng Named Ranges (Đặt tên vùng ô)
var namedRangeManager = new NamedRangeManager();
namedRangeManager.Define("GiaBan", "Sheet1!$C$2:$C$100");

// Evaluate công thức sử dụng tên đã đặt
var result = sheet.EvaluateFormula("SUM(GiaBan)");

// 3. Tham chiếu 3D (Tính tổng ô A1 từ Sheet1 đến Sheet3)
sheet["B1"].SetFormula("SUM(Sheet1:Sheet3!A1)");
```

### 5. Quản lý Bố cục & Hiển thị (Layout & Visibility)

```csharp
// Gộp ô (Merge)
sheet.MergeCells("A1:H1");

// Cố định dòng 1 và cột 1 (Freeze Panes)
sheet.FreezePanes(1, 1);

// Chèn bộ lọc tự động (AutoFilter)
sheet.SetAutoFilter("A1:D1");

// Ẩn/Hiện Worksheet
sheet.SetSheetVisibility(DocumentFormat.OpenXml.Spreadsheet.SheetStateValues.Hidden);

// Đặt độ rộng cột và chiều cao dòng thủ công
sheet.SetColumnWidth(1, 30.5); // Cột A rộng 30.5
sheet.SetRowHeight(1, 40);     // Dòng 1 cao 40
```

---

## ⚡ Tối ưu Hiệu năng

Thư viện được xây dựng với tư duy "Performance First":
*   **Shared String Table**: Tự động quản lý bảng chuỗi dùng chung để giảm kích thước file.
*   **Style Deduplication**: Tự động phát hiện và gộp các định dạng giống nhau, tránh phình to Stylesheet.
*   **LRU AST Cache**: Cache cây cú pháp công thức (AST) giúp việc tính toán hàng vạn dòng cực nhanh.
*   **Array Pool**: Tái sử dụng mảng trong quá trình tính toán để giảm thiểu Garbage Collection (GC).

---

## ⚠️ Lưu ý sử dụng

1.  **Chỉ số**: Row và Column bắt đầu từ **1** (theo chuẩn Excel).
2.  **IDisposable**: Luôn sử dụng `using` hoặc gọi `.Dispose()` cho `Workbook` để giải phóng bộ nhớ.
3.  **Lưu file**: Sau khi gọi `SaveAs`, workbook sẽ tự động được giải phóng để bảo mật dữ liệu và tài nguyên.

---

## 🗺️ Roadmap (Lộ trình phát triển)

*   [ ] Hỗ trợ Spill / Dynamic Array (Excel 365).
*   [ ] Pivot Tables (Bảng tổng hợp).
*   [ ] Digital Signatures (Chữ ký số).
*   [ ] Hỗ trợ thêm các hàm tài chính phức tạp.

---

© 2026 TVE Open Source Team. Phát hành dưới giấy phép MIT.
