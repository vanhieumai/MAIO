# Tối ưu frmExpandRows

## Nguồn và phạm vi
Đọc trực tiếp VBA từ D:\OneDrive\MyProject\CPMS\CPMS AR Plus.xlsm (bản lưu ngày 29/09/2026, 13:43). Chỉ sửa **UserForm frmExpandRows**. Không sửa modExportReport, Library, Reminder hoặc Designer. Workbook gốc chưa được chỉnh trực tiếp.

## Điểm gây chậm trong source
- MeasureWrapHeight tạo/xóa Shape TextBox mỗi lần đo một ô, và gọi DoEvents hai lần cho mỗi lần đo.
- Quét từng ô qua Excel object model, kể cả ô trống trong UsedRange.
- ExpandRowHeights tạo rồi xóa một sheet tạm nhưng luồng đo hiện tại không dùng sheet đó.
- ReDim Preserve ba mảng cho từng vùng merge nhiều dòng.
- Gán RowHeight từng dòng dù rất nhiều dòng có cùng chiều cao.

## Thay đổi
1. Dùng lại một Shape đo cho mỗi sheet; đặt lại AutoSize, Width, Height và font trước lần đo tiếp theo, xóa ở cuối hoặc khi lỗi.
2. Cache tối đa 2.000 kết quả/sheet theo text, chiều rộng, tên/cỡ font, Bold và Italic — đúng các thuộc tính mà thuật toán cũ sử dụng.
3. Đọc giá trị theo khối tối đa khoảng 65.536 ô và 256 dòng; bỏ qua ô trống/ô lỗi trước khi đọc định dạng và MergeArea. Vẫn duyệt theo thứ tự dòng/cột cũ.
4. Bỏ sheet tạm không sử dụng. Tăng dung lượng mảng merge theo cấp đôi.
5. Gộp các dòng liên tiếp cùng chiều cao để gán RowHeight một lần cho cả nhóm.
6. Chỉ nhường xử lý giao diện theo nhịp khoảng 0,25 giây, không DoEvents hai lần/ô; hiển thị sheet và dòng đang chạy trên StatusBar.
7. Chặn Start/Browse/đóng form khi đang chạy; khôi phục Calculation, EnableEvents, ScreenUpdating, DisplayAlerts, StatusBar và EnableCancelKey đúng giá trị ban đầu, kể cả khi lỗi. Hộp hoàn tất có thời gian chạy.

Giữ các công thức ước lượng, NormalizeWrapText, khoảng đệm, độ rộng trừ hao, phân bổ merge nhiều dòng và trần 409,5 pt. Vẫn xử lý phạm vi UsedRange như bản cũ, không tự ý thu hẹp vùng định dạng. UsedRange rất lớn do định dạng thừa vẫn có thể tốn thời gian.

## Cài đặt
1. Sao lưu chương trình và một workbook mẫu để thử.
2. Đóng form Expand Rows; Alt+F11 → Forms → frmExpandRows → F7.
3. Thay toàn bộ code bằng **frmExpandRows_FULL_CODE.txt** trong thư mục này. Không Import TXT hoặc xóa Designer.
4. Debug → Compile VBAProject; lưu chương trình, mở lại form và thử.

## Kiểm chứng và thử nghiệm thực tế
Đã kiểm tra 1.000 trường hợp gộp các dòng để đảm bảo chiều cao từng dòng giống phép gán riêng; kiểm tra source các hàm ước lượng giữ nguyên, reset Shape, cleanup, phục hồi trạng thái và không trùng tên thủ tục.

Chưa chạy VBA hoặc đo tốc độ trực tiếp trong Excel; không khẳng định nhanh hơn bao nhiêu lần. Việc bỏ DoEvents trong phép đo và dùng lại Shape cần được kiểm tra với phiên bản Excel thực tế: chạy bản cũ và bản mới trên hai bản sao cùng workbook, cùng MinHeight, so sánh thời gian và chiều cao, đặc biệt tiếng Việt, nhiều xuống dòng, font khác nhau, merge ngang/dọc, các ô rộng/hẹp xen kẽ, sheet rỗng và sheet bảo vệ.

Giữ cách lưu cũ: workbook do Browse mở mới được tự Save sau khi chạy thành công; workbook đã mở sẵn cần người dùng lưu. Khi gặp lỗi, chương trình dừng và báo sheet/lỗi; những chiều cao đã áp ở các sheet trước được giữ, không có rollback toàn workbook. Việc dọn Shape đã được gọi khi lỗi; nếu sheet/workbook bị đóng hoặc bảo vệ giữa chừng, cần kiểm tra lại đối tượng tạm.
