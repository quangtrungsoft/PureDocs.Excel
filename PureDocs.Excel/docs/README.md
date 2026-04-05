# TVE.PureDocs.Excel - Hướng dẫn sử dụng

Thư viện **TVE.PureDocs.Excel** cung cấp một bộ công cụ mạnh mẽ để làm việc với tệp Excel (.xlsx) trong .NET. Thư viện này hỗ trợ tạo, đọc, ghi, định dạng và tính toán công thức Excel với hiệu suất cao.

## 📋 Danh sách Chức năng chính

1.  **Quản lý Workbook & Worksheet**:
    *   Tạo mới (`Create`), Mở (`Open`) và Lưu (`SaveAs`) workbook từ File hoặc Stream.
    *   Quản lý danh sách Worksheet: Thêm mới, truy cập theo tên hoặc index.
2.  **Thao tác với Cell (Ô tính)**:
    *   Đọc/Ghi giá trị đa dạng: Text, Number, DateTime, Boolean.
    *   Tự động quản lý SharedStingTable để tối ưu dung lượng file.
    *   Hỗ trợ Cell Indexer (ví dụ: `sheet["A1"]` hoặc `sheet[1, 1]`).
3.  **Engine Tính toán Công thức (Formula Engine)**:
    *   Hỗ trợ gán công thức cho ô (`SetFormula`).
    *   **Smart Recalculation**: Tính toán lại thông minh dựa trên cây phụ thuộc (chỉ tính lại những ô bị ảnh hưởng).
    *   Hỗ trợ hàng loạt hàm Excel thông dụng: Toán học, Logic, Văn bản, Ngày tháng, Thống kê, Tra cứu (Lookup).
4.  **Định dạng & Styling (Fluent API)**:
    *   Hỗ trợ API dạng chuỗi (fluent) giúp viết code gọn gàng.
    *   Định dạng Font: Bold, Italic, Size, Color, FontName.
    *   Định dạng Fill: Background Color.
    *   Định dạng Border: Kẻ khung viền.
    *   Định dạng Alignment: Căn lề ngang/dọc, Wrap Text.
    *   Định dạng Số (Number Format): Date, Currency, Custom format.
5.  **Bố cục (Layout)**:
    *   Merge/Unmerge cells.
    *   Freeze Panes (đóng băng dòng/cột).
    *   AutoFit Column (tự động giãn cột theo nội dung).
    *   Set Row Height, Column Width.

---

## 🚀 Hướng dẫn sử dụng & Ví dụ

### 1. Khởi tạo và Lưu Workbook cơ bản

Dưới đây là ví dụ tạo một file Excel mới, thêm dữ liệu và lưu lại.

```csharp
using TVE.PureDocs.Excel;

// Tạo workbook mới
using var workbook = Workbook.Create();

// Thêm worksheet mới
var sheet = workbook.AddWorksheet("SalesData");

// Ghi dữ liệu vào ô
sheet["A1"].SetValue("Product ID");
sheet["B1"].SetValue("Product Name");
sheet["C1"].SetValue("Price");
sheet["D1"].SetValue("Quantity");
sheet["E1"].SetValue("Total");

sheet["A2"].SetValue(101);
sheet["B2"].SetValue("Laptop Dell");
sheet["C2"].SetValue(1500.00);
sheet["D2"].SetValue(2);

// Lưu file
workbook.SaveAs("SalesReport.xlsx");
```

### 2. Định dạng Cell (Styling)

Sử dụng Fluent API để định dạng ô nhanh chóng.

```csharp
// Định dạng Header (In đậm, nền xám, căn giữa, kẻ viền)
sheet.GetRange("A1:E1").Cells.ForEach(cell => 
{
    cell.Style
        .SetBold(true)
        .SetBackgroundColor(ExcelColor.LightGray)
        .SetHorizontalAlignment(ExcelHorizontalAlignment.Center)
        .SetAllBorders(ExcelBorderStyle.Thin);
});

// Định dạng cột Giá tiền (Currency)
sheet["C2"].SetNumberFormat("#,##0.00 $");

// Định dạng ngày tháng
sheet["F1"].SetValue(DateTime.Now);
sheet["F1"].Style.SetNumberFormat("dd/MM/yyyy");
```

### 3. Làm việc với Công thức (Formulas)

Thư viện hỗ trợ tính toán công thức ngay trong code.

```csharp
// Gán công thức tính tổng: Total = Price * Quantity
sheet["E2"].SetFormula("=C2*D2");

// Tính toán lại tất cả công thức trong sheet
sheet.CalculateFormulas();

// Lấy giá trị sau khi tính toán
var totalValue = sheet["E2"].GetValue(); // Kết quả: 3000.00
```

**Tính toán Recalculation thông minh (Hiệu năng cao):**

Nếu bạn thay đổi một giá trị, chỉ cần gọi `SmartRecalculate()`, thư viện sẽ chỉ tính lại các ô bị ảnh hưởng thay vì toàn bộ workbook.

```csharp
// Cập nhật số lượng
sheet["D2"].SetValue(5);

// Chỉ tính lại các ô phụ thuộc vào D2 (tức là E2)
sheet.SmartRecalculate();

var newTotal = sheet["E2"].GetValue(); // Kết quả: 7500.00
```

### 4. Các tính năng Layout & Hiển thị

```csharp
// Tự động giãn cột cho vừa nội dung
sheet.AutoFitColumn(2); // Cột B (Product Name)

// Đóng băng dòng tiêu đề (Freeze Header)
sheet.FreezePanes(1, 0); // Đóng băng dòng 1

// Merge cells (Gộp ô)
sheet["A10"].SetValue("Notes");
sheet.MergeCells("A10:E10"); // Gộp từ A10 đến E10
```

### 5. Đọc file Excel có sẵn

```csharp
using var workbook = Workbook.Open("ExistingReport.xlsx");
var sheet = workbook.Worksheets[0];

// Đọc giá trị
var productName = sheet["B2"].GetText();
var price = sheet["C2"].GetValue();

Console.WriteLine($"Product: {productName}, Price: {price}");
```

---

## 📚 Danh sách Nhóm hàm hỗ trợ (Formula Functions)

Dưới đây là các nhóm hàm được hỗ trợ trong Engine tính toán:

1.  **Math (Toán học)**: `SUM`, `ABS`, `ROUND`, `CEILING`, `FLOOR`, `POWER`, `SQRT`, ...
2.  **Logical (Logic)**: `IF`, `AND`, `OR`, `NOT`, `TRUE`, `FALSE`, `IFERROR`.
3.  **Text (Chuỗi)**: `LEFT`, `RIGHT`, `MID`, `LEN`, `TRIM`, `UPPER`, `LOWER`, `CONCATENATE`, `TEXT`.
4.  **Lookup (Tra cứu)**: `VLOOKUP`, `HLOOKUP`, `MATCH`, `INDEX`, `CHOOSE`.
5.  **Date & Time (Thời gian)**: `TODAY`, `NOW`, `DATE`, `DAY`, `MONTH`, `YEAR`.
6.  **Statistical (Thống kê)**: `AVERAGE`, `MIN`, `MAX`, `COUNT`, `COUNTA`, `COUNTIF`.

---

## ⚠️ Lưu ý

*   Chỉ số dòng (Row) và cột (Column) bắt đầu từ **1** (không phải 0).
*   Công thức hỗ trợ cú pháp tiếng Anh chuẩn (dấu phẩy `,` ngăn cách tham số).
*   `SmartRecalculate` tối ưu hơn `CalculateFormulas` khi cập nhật dữ liệu liên tục.
