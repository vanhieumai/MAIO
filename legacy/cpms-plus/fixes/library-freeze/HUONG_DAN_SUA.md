# Sửa treo Library sau khi thêm APP-3

## Vị trí sửa

Chỉ sửa **UserForm `frmContractSearch`** trong VBA project của **file chương trình CPMS AR.xlsm / add-in CPMS đang dùng**. Không sửa modLibrary, frmContractEdit, file data `.cpms` hoặc giao diện thiết kế UserForm.

## Lỗi đã tìm thấy

Thứ tự tab hiện tại (đếm từ 0): Information=0, Document=1, Library=2, Payment=3, Progress=4, Export=5.

Nhưng LoadContractDetail còn dùng các chỉ số cũ:

```vb
If tabMain.value = 2 Then LoadPaymentTab contractID
If tabMain.value = 3 Then UpdateExportPreview
If tabMain.value = 3 Then RefreshAllLibraryLabels
```

Khi một hợp đồng có giao dịch được chọn lúc đang ở Library, đường gọi là:

```text
LoadContractDetail
  → LoadPaymentTab                  (gọi nhầm trên Library)
    → tự chọn giao dịch / lstPayment_Click
      → SelectContract
        → LoadContractDetail
          → LoadPaymentTab ...      (lặp lại)
```

LoadPaymentTab vừa đổi ListIndex vừa gọi lstPayment_Click trực tiếp, nên đường gọi lặp tồn tại ngay cả khi không tính lần Click tự phát sinh của MSForms. Chỉ đổi `2` thành `3` sẽ chuyển vòng lặp sang tab Payment; phải tách chọn hợp đồng khỏi nạp lại danh sách thanh toán.

Data xác nhận B-11.3/APP-3 có ID B-2026-03-03. Ảnh treo đang chọn B-5 (ID B-2026-16), hợp đồng có giao dịch và đi được vào vòng gọi này. Chưa có call stack của phiên Excel treo để khẳng định đây là nguyên nhân duy nhất của phiên đó. Không thấy cần xóa APP-3 hoặc sửa data để xử lý vòng lặp đã xác định.

## Cách áp dụng đầy đủ

1. Tạo bản sao file chương trình trước khi thay code. Mở đúng file chương trình có frmContractSearch trong VBA project; file data không có UserForm này.
2. Nhấn Alt+F11, chọn **Forms → frmContractSearch → View Code (F7)**.
3. Mở `frmContractSearch_FULL_CODE.txt` trong cùng thư mục bản vá. Copy toàn bộ nội dung.
4. Trong cửa sổ **code của frmContractSearch**, Ctrl+A rồi dán thay toàn bộ code. Không xóa UserForm trong Project Explorer, không dùng Import File cho file TXT.
5. Chọn Debug → Compile VBAProject, rồi lưu file chương trình. Nếu có lỗi compile ở thành phần khác, ghi lại tên và dòng lỗi; bản vá này không sửa các lỗi cũ ngoài frmContractSearch.
6. Đóng/mở lại form và thử các tình huống bên dưới.

File FULL_CODE giữ nguyên các thủ tục ngoài sáu thủ tục được sửa. Nó được tạo từ bản CPMS AR.xlsm đã cung cấp. Nếu file đang sử dụng đã có thay đổi mới sau bản này, dùng cách sửa từng vị trí dưới đây để tránh ghi đè thay đổi mới.

## Cách sửa từng vị trí

### 1–2. Thay hai thủ tục

Trong frmContractSearch, thay toàn bộ `LoadContractDetail` và `SelectContract` bằng hai thủ tục trong `replacement_procedures.txt`. Tham số mới `refreshActiveTab` có giá trị mặc định True nên các lời gọi cũ một tham số vẫn hợp lệ.

### 3. Trong lstPayment_Click

Đổi:

```vb
SelectContract contractID
```

thành:

```vb
SelectContract contractID, False
```

Đây là bước ngắt vòng lặp: chọn giao dịch vẫn cập nhật hợp đồng và phần thông tin, nhưng không nạp lại danh sách thanh toán vừa được chọn.

### 4. Trong lstResult_Click

Đổi:

```vb
LoadContractDetail contractID
```

thành:

```vb
LoadContractDetail contractID, False
```

Giữ nguyên các lệnh xử lý tab phía dưới handler. Chúng đã dùng đúng index và sẽ nạp tab đúng một lần.

### 5. Trong LoadPaymentTab

Thêm ngay sau dòng khai báo Sub, trước `lstPayment.Clear`:

```vb
Erase mPayRowContractID
```

Mảng này là ánh xạ dòng khi xem All A/B. Nếu không xóa khi dựng lại danh sách riêng một hợp đồng, ResolvePayRowContractID có thể lấy ID cũ của danh sách trước. Nhánh All A/B sẽ dựng lại mảng riêng như code hiện có.

### 6. Trong LoadContractList

Trong khối `If autoContractID <> "" Then`, giữ:

```vb
LoadContractDetail autoContractID
If tabMain.value = 1 Then LoadDocumentList autoContractID
```

Xóa khối cũ:

```vb
If tabMain.value = 4 Then
    RefreshLibraryMeta
    RefreshAllLibraryLabels
End If
```

Index 4 là Progress. LoadContractDetail mới đã xử lý Library tại index 2.

## Kiểm tra sau khi dán

- Chọn B-11.3/APP-3 rồi vào Library; đổi qua B-11, B-11.1, B-11.2 và B-5 khi Library đang mở.
- Mở Payment của B-5 hoặc một hợp đồng có nhiều giao dịch. Click từng giao dịch; danh sách không được tự nạp lặp hoặc mất lựa chọn.
- Bật All A/B, chọn giao dịch, tắt All A/B rồi chọn giao dịch trong danh sách riêng: hợp đồng không được nhảy sang ID của danh sách cũ.
- Sau thêm/sửa hợp đồng, kiểm tra hợp đồng được chọn và Library hiển thị metadata/số hồ sơ đúng.
- Chuyển Document, Progress, Export để kiểm tra điều hướng vẫn đúng.

## Đã kiểm tra trong môi trường đọc source

6 kiểm tra đạt: xác nhận ID từ CHC1; mô hình callback code cũ tái hiện vòng lặp ở Library; mô hình bản sửa không nạp Payment từ Library; chọn giao dịch ở Payment kết thúc hữu hạn; các điểm gọi truyền cờ đúng; phạm vi thay đổi chỉ có sáu thủ tục trong một form.

Đây là kiểm tra source và mô hình callback, **chưa phải chạy VBA/Compile hoặc thử UserForm trong Excel**. File workbook và data gốc chưa bị thay đổi. Bản vá được cung cấp dưới dạng source để dán vào VBA Editor, giữ nguyên tài nguyên thiết kế form.
