# Reminder v3: giữ cố định Refresh và Go to

Chỉ sửa UserForm **frmReminder**. Thứ tự nút:

**Delete orphan | Refresh | Go to | Assign contract**

Refresh và Go to giữ nguyên vị trí trong Designer. Hai nút ngoài lấy khoảng cách giữa hai nút gốc làm khoảng cách đặt nút (tối thiểu 8 pt). Kích thước, font và cơ chế chỉ hiện khi chọn document mất hợp đồng giữ nguyên như v2.

## Cài đặt
Đóng Reminder, Alt+F11 → frmReminder → F7. Thay toàn bộ code bằng **frmReminder_FULL_CODE.txt** trong thư mục này. Debug → Compile VBAProject, lưu và mở lại form. Giữ nguyên Designer.

Nếu chỉ muốn sửa ít code từ bản v2, thay đúng ba thủ tục trong **replacement_procedures.txt**: CreateOrphanDeleteButton, StyleOrphanButtons, UpdateReminderActions. Hai biến mRefreshLeft và mGoToLeft cũ có thể xóa vì không còn dùng.

Bản này vẫn gồm đầy đủ chức năng xóa và gán ContractID, xem trước thông tin hợp đồng, xác nhận và lưu dữ liệu của v2. Không cần sửa module khác.

Đã kiểm tra diff và xác nhận code không còn lệnh gán Left/Top cho Refresh hoặc Go to. Chưa chạy giao diện Excel; cần kiểm tra chuyển qua lại giữa document hợp lệ và document mất hợp đồng. Workbook gốc chưa được chỉnh trực tiếp.
