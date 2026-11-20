# Sửa Ctrl+click mã Library — frmContractSearch

## Phát hiện và giới hạn
Code hiện tại gọi InputBox trực tiếp từ lblValSubCon_MouseDown (và bốn nhãn mã còn lại). Hộp thoại được mở khi thao tác nhấn chuột chưa kết thúc. Đây là đường xử lý có nguy cơ giữ mouse capture/kẹt tương tác của form modeless sau khi đóng hộp thoại. SaveContractMetaField trong modGlobal chỉ tìm dòng, ghi cột Z và cập nhật cache; trong đường gọi này không thấy vòng lặp vô hạn hay lệnh dựng lại lstResult.

Chưa tái hiện lỗi trong Excel nên chưa thể khẳng định đây là nguyên nhân duy nhất của Not Responding. Bản vá xử lý điểm rủi ro cụ thể trên và bảo vệ trạng thái chọn hợp đồng trong lúc hộp nhập mở.

## Vị trí sửa
Chỉ UserForm **frmContractSearch**, không sửa frmReminder hay module khác.

- lblValSubCon_MouseDown được thay bằng **lblValSubCon_MouseUp**. Áp dụng tương tự cho ProjCode, Client, MainCon, SupCon.
- Thêm API LibReleaseCapture (ReleaseCapture của user32, PtrSafe giống môi trường VBA hiện tại).
- LibMeta_CtrlClick: giải phóng mouse capture trước InputBox, chặn gọi lặp, giữ nguyên trạng thái suppressListClick rồi khôi phục khi OK/Cancel/lỗi.
- Chụp ContractID và workbook trước hộp nhập; kiểm tra lại trước khi ghi để tránh gán mã sang hợp đồng/dự án khác.
- Nạp lại cache hợp đồng trước đọc mã SUBCON, tránh cache cũ sau tạo hợp đồng mới.
- lstResult_Click và tabMain_Change không xử lý lặp trong lúc chỉnh mã.

Bản này kế thừa bản sửa Library trước ở fixes/library-freeze, bao gồm sửa thứ tự tab và luồng chọn Payment. Không dựng lại lstResult khi lưu mã viết tắt. Cơ chế ghi dữ liệu giữ như cũ: SaveContractMetaField ghi CONTRACT_DB cột Z vào workbook dữ liệu đang mở; vẫn cần lưu dữ liệu bằng cơ chế hiện có của ứng dụng.

## Cài đặt
1. Sao lưu workbook chương trình; đóng form Project Management.
2. Alt+F11 → Forms → **frmContractSearch** → F7.
3. Thay toàn bộ code bằng **frmContractSearch_FULL_CODE.txt** trong thư mục này. Giữ nguyên Designer; không Import TXT thành form mới.
4. Debug → Compile VBAProject, lưu workbook rồi mở lại form.

Nếu bạn đã sửa frmContractSearch thêm ngoài bản code đã cung cấp trong cuộc trò chuyện, đối chiếu frmContractSearch.diff trước khi thay toàn bộ để giữ những thay đổi riêng đó.

## Thử lại trong Excel
- Tạo contract mới, chọn Library, Ctrl+click Subcontractor (giữ Ctrl đến khi nhả chuột), nhập GTS và OK.
- Chọn hợp đồng khác trong lstResult, quay lại hợp đồng vừa sửa: mã phải đúng, danh sách vẫn tương tác được.
- Thử cả danh sách đầy đủ và kết quả tìm kiếm chỉ một dòng.
- Thử Cancel, mã không đổi và xóa mã bằng chuỗi trống.
- Thử các nhãn mã chung để kiểm tra tương tác sau đóng hộp nhập.
- Lưu dữ liệu, đóng/mở lại để kiểm tra mã đã lưu.

Đã kiểm tra tĩnh: không còn hộp nhập từ MouseDown của năm nhãn; ReleaseCapture trước InputBox; Cancel/lỗi khôi phục cờ; kiểm tra workbook/ContractID trước ghi; không thêm DoEvents hoặc LoadContractList trong handler; không trùng tên thủ tục. Chưa Compile hoặc chạy Excel thực tế. Workbook gốc chưa được chỉnh trực tiếp.
