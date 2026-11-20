# Library load v2 — frmContractSearch

## Phát hiện
Bản Ctrl+click trước chưa khắc phục hiện tượng vào Library bị treo. Kiểm tra thêm source cho thấy RefreshAllLibraryLabels gọi CountDocTrackerByType 25 lần, mỗi lần đọc từng ô từ DOCUMENT_TRACKER; GetTemplatePath cũng quét LIBRARY_DB lặp lại cho từng nhãn. LoadLibraryPage bỏ qua mọi lỗi và không có cờ chống nạp lại.

Đây là các vấn đề xác định được trong code. Chưa có stack hoặc nhật ký từ Excel đang treo, nên chưa thể khẳng định chúng là nguyên nhân duy nhất. Không xem bản này là đã xác nhận hết lỗi trên máy người dùng.

## Thay đổi
Chỉ thay **UserForm frmContractSearch**:
- ReadLibrarySnapshot đọc DOCUMENT_TRACKER A:C một lần vào mảng, đếm tất cả loại document trong một lượt; LIBRARY_DB cũng đọc một lần. Không dùng cache cũ nên thay đổi document được phản ánh ở lần nạp tiếp theo.
- Giữ quy tắc lọc ContractID, mã loại và alias IPC/THANH TOAN như luồng cũ.
- LoadLibraryPage, RefreshAllLibraryLabels và lstResult_Click có cờ chống gọi lặp, trả cờ khi lỗi/thoát.
- Không dùng DoEvents trong các thủ tục nạp trên và không dựng lại lstResult để vẽ Library.
- Lỗi nạp Library được báo, không bị nuốt im lặng.
- Thêm nhật ký từng bước, gồm đổi tab, chọn dòng, nạp detail, metadata, đọc hai sheet và cập nhật từng nhãn. Nhật ký đóng file sau mỗi dòng để còn dấu vết nếu Excel treo.

Kế thừa các bản sửa Library và Ctrl+click trước. Không thay frmReminder, modLibrary hoặc dữ liệu gốc.

## Cài đặt
1. Sao lưu chương trình. Đóng form Project Management.
2. Alt+F11 → Forms → frmContractSearch → F7.
3. Thay toàn bộ code bằng frmContractSearch_FULL_CODE.txt trong thư mục library-load-v2. Giữ nguyên Designer.
4. Debug → Compile VBAProject, lưu workbook, mở lại form.

Bản cuối đã đối chiếu và dựa trực tiếp trên D:\OneDrive\MyProject\CPMS\CPMS AR Plus.xlsm, lưu ngày 29/09/2026 09:12:19; SHA256 bcdafe89315da5c8b763e81d5f54b892f003bfac833d1026373bd1408b6b7b8a. File này có bản sửa Library đầu tiên nhưng chưa có bản Ctrl+click. Các thủ tục ngoài phạm vi sửa được giữ nguyên từ workbook hiện tại. Nếu đang có chỉnh sửa VBA chưa lưu trong Excel, hãy lưu bản riêng và đối chiếu trước khi thay code.

## Thử và lấy log
Thử chọn B-33, vào Library; chuyển sang B-39 và các hợp đồng khác; thử cả tìm kiếm một dòng; thử Ctrl+click sửa mã rồi Cancel/OK. Số document phải đúng và lstResult vẫn chọn được.

Nếu vẫn treo: mở Win+R, nhập **%TEMP%**, lấy file **CPMS_Library_trace.log** gửi lại cùng workbook/source hiện đang chạy. Dòng cuối cho biết bước đã bắt đầu trước khi dừng; nó là dấu vết chẩn đoán, không tự chứng minh nguyên nhân. File được giới hạn khoảng 256 KB rồi ghi vòng lại. Không cần lặp lại nhiều lần nếu đã treo.

Nhật ký chứa thời gian, bước chạy, số dòng, tên nhãn và thông báo lỗi; không ghi nội dung hợp đồng/document. Nếu file log không ghi được, ứng dụng bỏ qua lỗi log và tiếp tục.

## Kiểm chứng
Đã so sánh số đếm cho mọi ContractID có trong snapshot CHC1 và 25 loại document với quy tắc cũ; kiểm tra source không còn quét sheet theo từng nhãn, các cờ được trả và không trùng tên thủ tục. Chưa Compile hoặc chạy VBA thực tế trong Excel. Workbook gốc chưa chỉnh trực tiếp.

