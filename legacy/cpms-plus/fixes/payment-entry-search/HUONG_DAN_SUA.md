# txtSearch trong frmPaymentEntry

Chỉ sửa **frmPaymentEntry**. Giữ TextBox tên **txtSearch** bạn đã thêm trong Designer; đặt Tag trống, MultiLine=False. Không cần thêm module hoặc class. Dùng hàm VN_NoSign đã có trong project, giống frmContractSearch.

## Cài
Đóng Payment Entry. Alt+F11 → frmPaymentEntry → F7 → thay code bằng frmPaymentEntry_FULL_CODE.txt → Debug → Compile VBAProject → lưu, mở lại form.

Nếu form trong add-in có sửa riêng khác với workbook CPMS AR Plus.xlsm đã cung cấp, dùng PATCH_PROCEDURES.txt: thêm bốn khai báo ở đầu form, thay hai thủ tục LoadPaymentHistoryFast và InitPlaceholders, thêm các hàm/sự kiện tìm kiếm mới. Không dán các dòng hướng dẫn tiếng Anh trong file patch thành code. Không để trùng tên thủ tục.

## Cách dùng
- Gõ từ khóa và Enter để tìm, giống txtKeyword; không tự lọc theo từng phím.
- Không phân biệt hoa/thường, dấu tiếng Việt; tìm chuỗi con trong số thanh toán, Payment Reference, các số tiền hiển thị, ngày dd/mm/yyyy và ID thanh toán.
- Esc hoặc nhấp phải hai lần trong 0,5 giây để xóa từ khóa, hiện lại toàn bộ khoản thanh toán của hợp đồng đang chọn.
- Đổi hợp đồng sẽ xóa bộ lọc. Nạp lại cùng hợp đồng giữ từ khóa đã áp dụng.
- Không có kết quả thì danh sách trống. Tooltip của txtSearch có số kết quả.
- Tổng trên đầu bảng vẫn là tổng của hợp đồng; không chuyển thành tổng kết quả lọc.
- Mỗi lần tìm sẽ bỏ chọn khoản thanh toán cũ và reset phần nhập giống khi nạp lịch sử, để không giữ ID sửa/xóa đã bị ẩn. Cột ẩn 0 vẫn chứa chỉ số ledger gốc; không dùng thứ tự dòng sau lọc để sửa/xóa.

Đã kiểm tra mô hình tìm có/không dấu, hoa/thường, ngày, tiền, rỗng/không kết quả và giữ mapping ledger. Chưa Compile/chạy trực tiếp trong Excel. Workbook gốc chưa sửa. Trên bản sao, thử chọn khoản thứ hai sau lọc và kiểm tra ID trước khi bật Edit Payment.
