# Sửa lỗi ExpandRowHeights: Invalid use of Null trong frmExpandRows

## 1. Hiện tượng lỗi
Khi bấm nút **Start** (`btnStart_Click`) trong form **frmExpandRows**, quá trình xử lý bị dừng ngay từ sheet đầu tiên (hoặc sheet cụ thể, ví dụ: `KH chuy?n vtu TRP`) và hiển thị thông báo:
```text
Stopped after 0 sheet(s). Changes already applied are retained.
ExpandRowHeights: KH chuy?n vtu TRP: Invalid use of Null
```

## 2. Nguyên nhân kỹ thuật
- Trong Excel VBA, khi một ô tính có chứa **Rich Text** (định dạng hỗn hợp giữa các ký tự trong cùng một ô, ví dụ một đoạn in đậm và đoạn in thường, hoặc các cỡ chữ khác nhau):
  - Thuộc tính `cell.Font.Size` trả về `Null`.
  - Thuộc tính `cell.Font.Bold` trả về `Null`.
  - Thuộc tính `cell.Font.Name`, `cell.Font.Italic` trả về `Null`.
  - Nếu một ô chứa giá trị `Null` (từ truy vấn/công thức/COM/database), `cell.Value` trả về `Null`.
- Trong hàm `ExpandRowHeights` và các hàm helper liên quan:
  - Dòng tính chiều cao 1 dòng: `oneLineH = CDbl(cellRng.Font.Size) * 1.5` -> Khi `cellRng.Font.Size` là `Null`, `CDbl(Null)` lập tức gây lỗi **Runtime Error 94: Invalid use of Null**.
  - Trong `MeasureWrapHeight`: `cacheKey = ... & CStr(fnt.Size) & CStr(fnt.Bold) ...` -> `CStr(Null)` gây lỗi.
  - Khi gọi hàm `EstimateTextWrapHeight(..., fnt.Size)` có chữ ký `ByVal fontSize As Double`, truyền `Null` vào tham số `Double` làm phát sinh lỗi `Invalid use of Null`.
  - Các phép ép kiểu `CStr(rawValue)`, `CStr(topCell.value)`, `CStr(cellRng.value)` nếu gặp `Null` đều phát sinh lỗi.
- Do hàm `ExpandRowHeights` sử dụng `On Error GoTo EH`, lỗi được bắt và ném lên `btnStart_Click` hiển thị hộp thoại cảnh báo như hình.

## 3. Các thay đổi đã thực hiện
1. Bổ sung các hàm helper an toàn chống `Null`:
   - `SafeStr(v)`: Chuyển đổi an toàn `Variant` sang `String`, trả về `""` nếu là `Null`, `Empty`, hoặc lỗi.
   - `SafeFontSize(fnt, defaultSize)`: Trả về cỡ chữ hợp lệ (mặc định 11pt) nếu `fnt.Size` là `Null`.
   - `SafeFontName(fnt, defaultName)`: Trả về tên font hợp lệ (mặc định "Arial") nếu `fnt.Name` là `Null`.
   - `SafeFontBold(fnt)`: Trả về `Boolean` (mặc định False) nếu `fnt.Bold` là `Null`.
   - `SafeFontItalic(fnt)`: Trả về `Boolean` (mặc định False) nếu `fnt.Italic` là `Null`.
2. Chuẩn hóa trong `MeasureWrapHeight`, `MeasureNonMergedCellHeight`, `EstimateNonMergedWrapHeight`, `CellNeedsMultipleLines`.
3. Trong vòng lặp chính của `ExpandRowHeights`:
   - Thay `CStr(rawValue)` bằng `SafeStr(rawValue)`.
   - Thay `CStr(topCell.value)` bằng `SafeStr(topCell.value)`.
   - Thay `CStr(cellRng.value)` bằng `SafeStr(cellRng.value)`.
   - Thay `oneLineH = CDbl(cellRng.Font.Size) * 1.5` bằng `oneLineH = SafeFontSize(cellRng.Font) * 1.5`.

## 4. Kiểm tra
- Đã vá và biên dịch thành công (Compile clean) vào workbook: `CPMS AR - Sign Date Sort Fix.xlsm`.
- File sao lưu trước khi vá: `CPMS AR - Sign Date Sort Fix.xlsm.bak`.
