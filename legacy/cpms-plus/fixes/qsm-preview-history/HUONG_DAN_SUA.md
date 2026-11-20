# QSM Toolkit: lịch sử cấu hình và Link Data trong Preview

## Những file cần cài
1. **Class Module mới `clsQSMLinkPlan`**: trong VBA Editor, File → Import File → chọn **clsQSMLinkPlan.cls**. Nếu class này đã có, thay code bằng clsQSMLinkPlan_CODE.txt, không tạo bản trùng tên.
2. **UserForm `frmQSMToolkiT`**: F7, thay toàn bộ code bằng **frmQSMToolkiT_FULL_CODE.txt**.
3. **UserForm `frmPreviewLink`**: F7, thay toàn bộ code bằng **frmPreviewLink_FULL_CODE.txt**.
4. Debug → Compile VBAProject, lưu chương trình. Giữ nguyên Designer; nút và combobox mới được tạo bằng code.

Đóng các form Toolkit/Preview và sao lưu workbook chương trình trước khi thay. Bản sửa dựa trên VBA đọc trực tiếp từ D:\OneDrive\MyProject\CPMS\CPMS AR Plus.xlsm. Workbook gốc chưa được sửa trực tiếp.

## Lịch sử
- Preview khi chưa chọn đủ sheet/cột đối chiếu/cột link, hoặc số cột hai bên không bằng nhau, sẽ mở danh sách tối đa **5 cấu hình gần nhất** để chọn bằng số.
- Khi mở Preview với cấu hình hợp lệ hoặc Link Data thành công, chương trình lưu tên/path workbook, sheet, cột đối chiếu và toàn bộ cột link được chọn ở cả hai listbox. Không lưu dữ liệu ô.
- Khôi phục sẽ điền lại hai combobox sheet, hai combobox cột đối chiếu, hai listbox, cùng thông tin Source/Target. Lịch sử gắn với đúng workbook gốc, không áp âm thầm vào một workbook khác đang active.
- Workbook nguồn và đích phải đang mở. Nếu chưa mở hoặc sheet không tồn tại, có thông báo; không tự đóng/mở file hoặc bỏ thay đổi chưa lưu.
- Với workbook đã có đường dẫn lưu, lịch sử tồn tại qua các phiên Excel. Với Book1/chưa lưu, lịch sử chỉ khôi phục trong phiên form hiện tại, khi đúng workbook gốc còn mở.
- Lần đầu cài chưa có lịch sử: chọn thủ công một lần rồi mở Preview. Không thể phục hồi những cấu hình chưa từng được ghi trước khi cài bản sửa.
- Lưu bằng VBA SaveSetting tại HKCU\Software\VB and VBA Program Settings\CPMS_QSMToolkit\LinkHistoryV1 trên máy hiện tại.

## Preview và giới hạn dòng
Hai combobox **Target: first row** và **Last row (inclusive)** giới hạn dòng trên sheet Target, tính cả hai đầu. Mặc định là dòng 1 đến dòng cuối có dữ liệu ở cột đối chiếu Target. Có thể chọn từ danh sách hoặc gõ số; với sheet trên 10.000 dòng, danh sách có hai mốc đầu/cuối và cho gõ số bất kỳ trong phạm vi.

Khi thay mốc và rời combobox/nhấn Tab, Preview được nạp lại. Khoảng dòng không hợp lệ làm trống preview và vô hiệu Link Data. Mỗi dòng preview ghi số dòng Target, tên đối chiếu, tối đa 4 giá trị nguồn và OK/Not Found. **Mọi cặp cột đã chọn đều được link**, không chỉ 4 cột hiển thị; phần Columns phía trên liệt kê toàn bộ mapping, có tooltip nếu dài.

Nút **Link Data** trong Preview hoạt động trên đúng cấu hình được chụp lúc mở Preview. Việc đổi combobox/listbox ở Toolkit sau đó không đổi nguồn/đích của cửa sổ Preview cũ. Double-click một dòng vẫn đi đến ô nguồn tương ứng.

Hai nút Link Data dùng chung class xử lý:
- Nút bên ngoài: toàn bộ dòng 1 đến dòng cuối như trước.
- Nút trong Preview: chỉ khoảng dòng đã chọn; các dòng ngoài khoảng được giữ nguyên.
- Có hộp xác nhận ghi rõ workbook/sheet, mapping cột và khoảng dòng trước khi ghi.
- Ghi công thức liên kết như nút cũ; ô không khớp hoặc nguồn trống/lỗi thành trống trong vùng được link. Nội dung cũ trong vùng này bị thay thế, không phải chế độ chỉ điền ô trống.
- Không tự Save workbook đích; người dùng kiểm tra rồi lưu.

## Sửa sai lệch cũ
Preview trước đây bắt đầu dò tên trùng từ dòng 10, trong khi Link Data bắt đầu từ dòng 1. Bản này thống nhất theo Link Data cũ: trim tên, phân biệt hoa/thường, ghép lần xuất hiện thứ n ở Target với lần thứ n ở Source, bắt đầu từ dòng 1.

Khi giới hạn Target từ dòng giữa, số lần xuất hiện trước dòng đầu vẫn được tính. Vì vậy thu hẹp khoảng dòng không làm tên trùng bị ghép lại với dòng nguồn đầu tiên.

Đọc các cột tên chỉ đến dòng cuối dữ liệu, không tải cả triệu dòng. Không bỏ filter trên sheet đích như code cũ; giới hạn dòng áp dụng theo số dòng thực, bao gồm cả dòng đang ẩn bởi filter.

## Kiểm chứng và giới hạn
12 kiểm tra mô hình/source đạt, gồm 500 trường hợp ngẫu nhiên kiểm tra giới hạn dòng không đổi kết quả ghép tên trùng; alias tên sheet có dấu nháy; giới hạn vùng ghi; lịch sử và nhận diện workbook chưa lưu; hai nút dùng chung engine; phục hồi trạng thái Excel.

Chưa Compile hoặc chạy trực tiếp trong Excel. Cần thử trên bản sao: lưu cấu hình → Reset → Preview → chọn lịch sử; kiểm tra đủ lựa chọn; Preview dòng 10–20 → Link → chỉ các dòng 10–20 thay đổi; thử tên trùng trước dòng 10, nguồn ngoài workbook, Cancel, nguồn đóng và đổi cấu hình ở form chính khi Preview vẫn mở.

Bản này báo lỗi trước khi ghi nếu đích read-only/protected, vùng đích có ô merge, hoặc đích trùng cột dữ liệu nguồn trên cùng sheet. Nếu lỗi ghi phát sinh giữa các cột, thông báo số cột đã ghi; không có rollback tự động. Không tự lưu dữ liệu. Với dữ liệu lớn, nên chọn phạm vi cần thiết để hạn chế bộ nhớ và thời gian dựng preview.
