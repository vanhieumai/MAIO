# Reminder v2 — ẩn/hiện thao tác và gán lại ContractID

## Nơi sửa
Chỉ thay code UserForm **frmReminder**. Không sửa module khác, không cần thêm control trong Designer. Code tự tạo các nút và khung chọn hợp đồng.

## Cài đặt
1. Sao lưu workbook chương trình .xlsm và dữ liệu .cpms, đóng Reminder.
2. Alt+F11 → Forms → frmReminder → F7.
3. Thay toàn bộ code hiện tại bằng nội dung **frmReminder_FULL_CODE.txt** trong thư mục reminder-orphan-v2. Không Import TXT thành form. Giữ nguyên Designer.
4. Debug → Compile VBAProject, lưu workbook, mở lại Reminder.

Bản này bao gồm chức năng Delete orphan của bản trước. Không thay frmContractSearch nếu đã sửa Library.

## Giao diện
- Delete orphan và Assign contract lấy chiều rộng, chiều cao, font, cỡ chữ, kiểu chữ và màu từ nút Refresh, đồng bộ lại sau ApplyTheme.
- Hai nút chỉ hiện khi document đang chọn có ContractID không tồn tại. Khi chọn document hợp lệ hoặc không chọn dòng, chúng ẩn đi.
- Nếu dữ liệu chỉ đọc hoặc sheet bị bảo vệ, hai nút hiện nhưng không cho thao tác.
- Khi hai nút hiện, Refresh và Go to được dịch sang phải thành một hàng. Khi chọn document bình thường, chúng trở về vị trí cũ.

## Gán hợp đồng
1. Chọn document có nhãn [Missing contract], bấm **Assign contract**.
2. Chọn hợp đồng/appendix trong danh sách. Danh sách hiển thị thứ tự, số hợp đồng/appendix và bên B.
3. Khung xem trước hiển thị DocumentID đang sửa, ContractID cũ, ContractID mới, thứ tự (ví dụ A-1.8), số hợp đồng/appendix (APP-17), loại, bên B, phạm vi công việc và trạng thái.
4. Bấm **Confirm assign**, kiểm tra thông tin trong hộp xác nhận rồi chọn Yes. Cancel hoặc No không ghi dữ liệu.

Chỉ DOCUMENT_TRACKER cột B của đúng DocumentID được đổi; số document, tên, trạng thái, ngày, file đính kèm và các dữ liệu khác giữ nguyên. Code không tự sửa mã document để khớp hợp đồng mới. Dữ liệu được lưu và cache/danh sách được nạp lại. Document vừa gán được giữ chọn nếu vẫn đáp ứng bộ lọc; hai nút sẽ ẩn vì document đã có hợp đồng.

Trước và sau xác nhận, code kiểm tra lại đúng workbook, DocumentID duy nhất, ContractID cũ vẫn mất, hợp đồng đích tồn tại duy nhất và thông tin đích không đổi. Không chọn dòng phân nhóm GROUP- làm hợp đồng.

## Các thủ tục chính
- CreateOrphanDeleteButton, MatchReminderButton, StyleOrphanButtons: tạo và đồng bộ nút.
- UpdateReminderActions: điều khiển Visible và Enabled.
- CreateContractPicker, mAssignContract_Click: mở danh sách hợp đồng.
- AssignmentContractRow, AssignmentContractInfo, mContractChoice_Change: kiểm tra và hiển thị thông tin hợp đồng.
- mConfirmAssign_Click: xác nhận, kiểm tra lại, đổi ContractID, lưu và cập nhật cache.
- CloseContractPicker, mCancelAssign_Click: đóng/hủy; thay document hoặc nạp lại danh sách cũng đóng phần chọn cũ.

## Kiểm tra
15 kiểm tra mô hình dữ liệu và source đã đạt: .vba-review/test_reminder_v2.py. Bao gồm chuyển IPC DOC-20260918-001 từ A-2026-02 sang A-2026-01-09 chỉ thay cột B; Cancel giữ nguyên; chặn nguồn hợp lệ/đích không tồn tại; kiểm tra lại sau xác nhận; hiển thị và kiểu nút; kiểm tra Delete orphan cũ.

Chưa chạy VBA/kiểm tra giao diện trực tiếp trong Excel. Trên bản sao hãy Compile, thử chọn dòng hợp lệ/mất hợp đồng/không có kết quả, kiểm tra hai nút không chồng nhau, mở Assign contract → chọn APP-17/A-1.8 → xem đủ thông tin → Cancel rồi thử Yes; đóng/mở lại dữ liệu để xác nhận lưu. Nếu không lưu được, thông báo phân biệt rõ thay đổi trong bộ nhớ với dữ liệu đã lưu.

Thao tác Save ghi cả thay đổi chưa lưu khác trong cùng workbook dữ liệu. Workbook gốc và file .cpms chưa được chỉnh trực tiếp bởi bộ bản vá này.
