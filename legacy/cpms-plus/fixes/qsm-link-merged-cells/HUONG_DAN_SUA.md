# Sửa lỗi Link Data khi cột điều kiện (cbCotTenGT, cbCotTenTH) hoặc cột đích có Merged Cells

## 1. Hiện tượng lỗi
Khi bấm nút **Link** (`btnLink_Click`) trên form **frmQSMToolkiT**, đôi khi xuất hiện thông báo lỗi:
```text
Cannot link: Target link range contains merged cells.
```
hoặc các dòng có ô tên bị merge không tìm thấy giá trị đối soát (Not Found / rỗng), khiến dữ liệu không được link sang hoặc bị dừng tiến trình.

## 2. Nguyên nhân kỹ thuật
1. **Ô điều kiện tìm kiếm bị merge**:
   - Khi một ô trong cột Tên (`cbCotTenTH` hoặc `cbCotTenGT`) bị merge theo chiều ngang (ví dụ merge từ B:C và người dùng chọn cột C) hoặc merge theo chiều dọc (ví dụ merge dòng 5:6):
   - Đọc bằng `Range(...).Value` khiến Excel trả về `Empty` cho mọi ô không phải ô top-left của vùng merge.
   - Hàm `End(xlUp)` khi tìm dòng cuối bị nhảy lên phía trên do các ô phía dưới bị rỗng trong cột phụ của vùng merge.
   - Kết quả: Từ khóa so sánh bị rỗng (`""`), dẫn đến không khớp được dòng nào hoặc tính sai phạm vi dòng (`LastTargetRow`).

2. **Cột đích (`mTargetCols`) hoặc tiêu đề bảng tính có ô merge**:
   - `btnLink_Click` thực hiện `plan.Execute(1, plan.LastTargetRow)` bắt đầu từ dòng 1.
   - Phần tiêu đề (header) của bảng tính thường có các ô merge tiêu đề cột (ví dụ dòng 1-2 merge).
   - Trước đây trong `clsQSMLinkPlan.Execute`, code kiểm tra:
     ```vba
     merged = targetRange.MergeCells
     If IsNull(merged) Then Err.Raise 5, , "Target link range contains merged cells."
     If merged Then Err.Raise 5, , "Target link range contains merged cells."
     ```
     và gán công thức bằng mảng khối: `targetRange.Formula = output`.
   - Vì có ô merge nên code chủ động ném lỗi và dừng hoàn toàn quá trình link data.

## 3. Các thay đổi đã thực hiện trong `clsQSMLinkPlan`

1. **Hàm `GetKeyColumnLastRow`**:
   - Quét ngược từ `UsedRange` để kiểm tra `MergeArea.Cells(1, 1).Value`, đảm bảo nhận diện chính xác dòng cuối cùng có dữ liệu ngay cả khi người dùng chọn cột phụ của ô merge ngang.
   - Cập nhật cả `LastTargetRow` và `lastSource` trong `BuildMatches`.

2. **Hàm `ReadKeyColumn`**:
   - Kiểm tra `hasMerge = rngCheck.MergeCells`.
   - Nếu có ô merge trong dải ô, duyệt và tự động lấy giá trị từ ô gốc:  
     `val = c.MergeArea.Cells(1, 1).Value`
   - Đảm bảo dù chọn cột nào hoặc dòng nào trong vùng merge ngang/dọc, từ khóa so sánh vẫn luôn lấy được đúng văn bản.

3. **Cập nhật `FillPreview` và tạo công thức `Execute`**:
   - Nếu ô nguồn dữ liệu bị merge, tự động trỏ về ô gốc `srcCell.MergeArea.Cells(1, 1)` để công thức link không bị nhận giá trị 0 hoặc rỗng.

4. **Xử lý ghi công thức thông minh khi có Merged Cells trong `Execute`**:
   - Loại bỏ lệnh ném lỗi chặn `Target link range contains merged cells.`
   - Nếu cột đích không có ô merge: Giữ nguyên cơ chế ghi công thức theo khối `targetRange.Formula = output` để đạt tốc độ tối đa.
   - Nếu cột đích có chứa ô merge: Chuyển sang cơ chế ghi an toàn từng dòng:
     + Dòng có match: Ghi vào `targetCell.MergeArea.Cells(1, 1).Formula = formulaVal`.
     + Dòng không match: Nếu ô là ô merge (ví dụ header dòng 1-3, ô nhóm) thì giữ nguyên, tránh lỗi `Cannot change part of a merged cell` của Excel. Chỉ clear nếu là ô đơn không merge.

## 4. Kiểm tra
- Đã vá và biên dịch thành công (Compile clean, không có lỗi cú pháp) trong workbook:
  `CPMS AR - Sign Date Sort Fix.xlsm`.
- Bản sao lưu trước khi vá: `CPMS AR - Sign Date Sort Fix.xlsm.bak_qsm`.
