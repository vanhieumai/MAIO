# Sửa Reminder: xóa document mất hợp đồng

Chỉ sửa UserForm **frmReminder**. Code tự tạo nút **Delete orphan**, không cần thêm control trong Designer hoặc sửa module khác.

## Nguyên nhân
Document liên kết bằng ContractID nội bộ trong DOCUMENT_TRACKER cột B, đối chiếu CONTRACT_DB cột A. Xóa appendix để lại document cũ; appendix tạo lại có ID khác dù cùng tên APP-17/A-1.8.

Trong dữ liệu CHC1, IPC DOC-20260918-001 giữ ID cũ A-2026-02, không còn tồn tại. APP-17 hiện tại mang ID A-2026-01-09; document mới DOC-20260925-002 được bảo vệ.

## Cài bản sửa
1. Sao lưu workbook chương trình .xlsm và dữ liệu .cpms. Đóng Reminder.
2. Alt+F11 → Forms → frmReminder → F7.
3. Mở frmReminder_FULL_CODE.txt, sao chép toàn bộ và thay toàn bộ code trong cửa sổ code của frmReminder. Giữ nguyên form và các control trong Designer. Không Import file TXT thành form mới.
4. Debug → Compile VBAProject, lưu .xlsm và mở lại Reminder.

Giữ nguyên frmContractSearch nếu đã cài bản sửa Library trước đó.

## Sử dụng
Chọn dòng [Missing contract] → Delete orphan → kiểm tra DocumentID, số document và ContractID → Yes. Chỉ một dòng trong DOCUMENT_TRACKER bị xóa; file tài liệu đính kèm không bị xóa.

Document còn ContractID hợp lệ không được xóa tại đây, kể cả hợp đồng Completed/Cancel hoặc thiếu tên công ty. Bật Show Completed để xử lý document orphan đã hoàn thành. Bộ lọc và tìm kiếm vẫn áp dụng.

Lưu dữ liệu cũng ghi các thay đổi chưa lưu khác trong cùng workbook. Nếu lưu thất bại, form báo rõ thao tác xóa mới xảy ra trong bộ nhớ, chưa xác nhận lưu xuống đĩa.

## Các phần code thay đổi
- mDeleteOrphan (WithEvents), mListWorkbook: nút và workbook nguồn.
- UserForm_Initialize, CreateOrphanDeleteButton: tạo nút.
- LoadPendingList: đánh dấu mất hợp đồng, lưu workbook nguồn; cấp phát buffer theo số dòng để tránh ReDim Preserve sai chiều khi danh sách lớn.
- lstPending_Click, UpdateReminderActions: bật/tắt Delete orphan và Go to.
- ReminderContractExists: kiểm tra trực tiếp CONTRACT_DB, không dựa vào cache.
- ResolveReminderSelection: tìm DocumentID duy nhất, đối chiếu ContractID và workbook, không dùng số dòng cũ.
- mDeleteOrphan_Click: xác nhận, kiểm tra lại, xóa, lưu và nạp lại cache/danh sách.
- btnGoTo_Click, CanNavigateReminderSelection: chặn điều hướng document mất hợp đồng.
- UserForm_Terminate: giải phóng tham chiếu bổ sung.

## Kiểm chứng
9 kiểm tra dữ liệu/mô hình và source đã đạt trong .vba-review/test_reminder_fix.py: IPC cũ; APP-17 mới; danh sách orphan; chỉ xóa dòng chọn; ID trùng/biến mất/đổi hợp đồng; chuyển workbook/khôi phục hợp đồng; kiểm tra lại sau xác nhận; cấu trúc source; hash workbook gốc.

Chưa chạy VBA hoặc kiểm tra giao diện trong Excel. Thử trên bản sao: Compile; IPC cũ → nút bật; IPC mới → nút tắt; No → giữ nguyên; Yes → mất đúng một dòng; đóng/mở lại dữ liệu → xác nhận đã lưu; Show Completed → kiểm tra orphan đã hoàn thành.

Workbook .xlsm và .cpms gốc chưa được chỉnh trực tiếp. frmReminder.diff và frmReminder.fixed.source dùng để đối chiếu.
