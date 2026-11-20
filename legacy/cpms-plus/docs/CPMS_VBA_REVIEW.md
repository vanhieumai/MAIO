# Ghi chú đọc source VBA CPMS AR

Ngày đọc: 25/09/2026. File nguồn: `CPMS AR.xlsm`.

Bổ sung sau khi đọc file data CHC1: xem `CPMS_CHC1_DATA_REVIEW.md`. PROJECT_SETTINGS có trong data dù thiếu trong workbook chương trình. Đã đối chiếu cây hợp đồng, giao dịch/ngoại tệ, Cancel và hồ sơ liên kết; không áp dụng nhận xét về template máy mới để kết luận data CHC1 thiếu sheet.

## Phạm vi và cách đọc

Đã trích xuất đủ 56 thành phần source VBA từ `xl/vbaProject.bin`: 27 standard modules, 19 UserForms, 10 class/document modules; tổng cộng 38.081 dòng, gồm cả comment, dòng trống và thuộc tính module. Bộ đếm khai báo nhận diện 1.136 Sub/Function/Property, không phải 1.136 chức năng độc lập.

Đã lập danh mục thủ tục toàn project, đọc sâu các luồng khởi tạo, quản lý dự án, session/cache, thêm/sửa/xóa hợp đồng, ghi thanh toán, sinh số và lưu hồ sơ, Reminder, xuất thanh toán, AI; đồng thời khảo sát các điểm vào của theme và QSM Toolkit. Chưa kiểm chứng từng nhánh giao diện, từng thuật toán đo chiều cao hoặc toàn bộ thao tác merge/theme bằng thực thi.

Đây là phân tích tĩnh. Không mở workbook bằng Excel, không chạy macro, không gọi API của ứng dụng, không thay đổi file Excel. Các vị trí dưới đây trỏ tới source đã trích xuất trong `.vba-review/source/`. File `.frm` ở đây là phần mã, không phải bản export đầy đủ kèm tài nguyên `.frx` để import lại giao diện.

## Kiến trúc đã xác định

| Nhóm | Thành phần chính | Vai trò |
| --- | --- | --- |
| Điểm vào | `modRibbon` và `customUI14.xml` | Nút CPMS gọi `RibbonCurrent`; menu mở Project Manager, tạo/import dự án; các nút Reminder, QSM, License, Ask AI, Theme |
| Khởi tạo và dự án | `modCPMSBootstrap`, `frmHomePage`, `frmProjectInfo` | Tạo template, danh mục dự án, tạo/clone/rename/import/export và chọn dự án hiện tại |
| Truy cập dữ liệu | `modDataSession` | Mở workbook dữ liệu ẩn, đếm tham chiếu, lấy sheet, lưu/đóng session |
| Bộ nhớ đệm | `modGlobal` | Nạp mảng và Dictionary cho hợp đồng, thanh toán, ledger, hồ sơ; đánh dấu project hiện tại |
| Hợp đồng | `modContract`, `modLinkEngine`, `frmContractEdit`, `frmContractSearch` | Nhóm A/B, hợp đồng/phụ lục/package, kiểm tra trùng, đồng bộ ledger, tra cứu và đánh số hiển thị |
| Thanh toán | `modPayment`, `frmPaymentEntry`, `frmPaymentSummary` | Ghi giao dịch, tổng hợp cha/con, quy đổi tiền, bộ lọc thời gian, chi phí nội bộ B-CPB |
| Hồ sơ | `frmDocumentEdit`, `frmReminder`, `modLibrary`, `modBuildPackage` | Theo dõi trạng thái, sinh số/phiên bản, đường dẫn tài liệu, mẫu và đóng gói ZIP |
| Báo cáo | `modExportReport` | Thông tin hợp đồng, hồ sơ, thanh toán, bảng tổng hợp |
| AI | `frmAskAI`, `modAISession` | Chọn nhà cung cấp/model, gửi text/ảnh, lưu key và ghi chú trong dữ liệu AI riêng |
| Theme | `modThemes`, `modThemeINI`, `modThemeFile`, `modThemeZip`, cache/resource/scanner và `frmThemeManager` | Màu, ảnh, manifest, cài/đóng gói theme, cập nhật form đang mở |
| Tiện ích Excel | `frmQSMToolkiT`, `frmMergeSheets`, `frmExpandRows`, `frmExportnewSlx` | Liên kết dữ liệu theo tên và thứ tự trùng, gộp sheet, xử lý name/link, giãn dòng và xuất sheet |
| License/ảnh | `modLicense`, `frmLicense`, `GenerateQR_ToCell`, `modFirebase`, `modAssets`, `modImageUtils` | License theo máy, QR và tài nguyên ảnh; hàm Firebase tồn tại riêng |

### Nơi lưu dữ liệu

`GetAppFolder` dùng `%LOCALAPPDATA%\CPMS_R\`. Hệ thống dự kiến có:

- `DataTemplate\CPMS_DataTemplate.xlsb` làm mẫu.
- `ProjectList.xlsb` quản lý danh sách dự án.
- `CurrentProject.ini` giữ ProjectID, tên, thư mục và DataFile đang chọn.
- `Projects\<ProjectID>\CPMS_Data.xlsb` chứa dữ liệu từng dự án.
- `AI\AI_Data.xlsb` giữ SETTINGS và AI_NOTES riêng cho AI.

Export `.cpms` thực hiện bằng sao chép file dữ liệu sang đuôi mới; không phải định dạng cơ sở dữ liệu riêng.

Workbook nguồn hiện có 8 sheet: CPMS, SYS_INFO, CONTRACT_DB, PAYMENT_LEDGER, DOCUMENT_TRACKER, SETTINGS, AI_NOTES, LIBRARY_DB. Ngoài CPMS, các sheet đều veryHidden trong cấu trúc workbook đã đọc.

### Quy ước nghiệp vụ

- Nhóm A tương ứng hợp đồng với chủ đầu tư (Main/EPC); nhóm B tương ứng SUB. Dòng `GROUP-A`/`GROUP-B` được đặt xen trong bảng.
- ID gốc sinh dạng `A-<năm>-01`/`B-<năm>-01`; phụ lục nối thêm đoạn `-01`. Số hiển thị APP-/PKG- là một khái niệm khác ID nội bộ.
- CONTRACT_DB đọc A:Z: A là ID, B số hợp đồng, D ngày ký, F trạng thái, O giá trị, R category; X/Y giữ thứ tự hiển thị/sort, Z giữ chuỗi metadata.
- PAYMENT_LEDGER dùng A làm cha, B làm LedgerID, C làm ContractID. Dòng giao dịch có B khác rỗng; dòng header có B rỗng. J:M chứa thực hiện/nghiệm thu/đề nghị/đã trả; N là chênh lệch đề nghị trừ đã trả. W:Z là các giá trị quy đổi, header tổng hợp bằng SUMIFS theo ParentID.
- DOCUMENT_TRACKER dùng A làm DocumentID, B làm ContractID, C loại, D số, F đường dẫn, I/J ngày, O trạng thái, Q/R bên phát hành/nhận, S/T subtype. Cache hiện chỉ đọc tới S.
- Số hồ sơ ghép ngày, mã dự án, bên phát hành, loại/subtype, bên nhận, số thứ tự và hậu tố phiên bản. Bộ đếm được lọc theo hợp đồng, loại và bên phát hành/nhận chính.

## Những điểm cần xử lý trước

### 1. Template yêu cầu một sheet không có trong file nguồn

`modCPMSBootstrap.bas:855` liệt kê PROJECT_SETTINGS để copy từ ThisWorkbook. `BuildDataTemplateWithProgress` kiểm tra sheet này tại dòng 678 và báo thiếu tại dòng 690. Workbook được cung cấp không có PROJECT_SETTINGS.

Hệ quả: đường khởi tạo template từ file này trên máy mới sẽ thiếu sheet bắt buộc. Máy đã có template cũ có thể không gặp ngay vì có nhánh bỏ qua theo version. Ngoài ra, nhánh báo lỗi tại dòng 733 nằm sau `Resume Cleanup`, trong khi Cleanup kết thúc bằng `Exit Sub`; thông báo này không được chạy theo luồng đó. Cần trả trạng thái thất bại rõ ràng trước khi đánh dấu version đã được tạo.

### 2. Chuyển dự án có thể bỏ thay đổi chưa lưu

`frmPaymentEntry.frm:154` và `:393` ghi trực tiếp vào sheet, refresh cache và báo thành công nhưng không gọi SaveDataFile. Lưu thông thường phụ thuộc CloseSession của frmContractSearch.

Trong khi đó, `frmHomePage.frm:421` đóng workbook cũ bằng `Close False` trước khi unload frmContractSearch; `modCPMSBootstrap.bas:1037` cũng đóng workbook dữ liệu khác bằng `Close False`. Nếu người dùng thêm/sửa thanh toán rồi chuyển dự án khi session cũ còn mở, phần thay đổi chưa lưu có nguy cơ mất. Chưa thực thi tình huống này trong Excel.

Chỉ frmContractSearch gọi OpenSession/CloseSession trong source tìm thấy; các form dữ liệu khác không thực sự tự giữ reference như comment trong modDataSession mô tả. Cần thống nhất chủ sở hữu session và lưu trước khi chuyển project.

### 3. Cập nhật thanh toán ghi dữ liệu trước khi validation hoàn tất

`frmPaymentEntry.frm:426` đã ghi tỷ giá, số tiền và tham chiếu trước khi kiểm tra ngày tại dòng 444. Nếu ngày sai, hàm Exit Sub nhưng các ô trước đó vẫn thay đổi. Người dùng có thể nghĩ cập nhật bị từ chối toàn bộ.

Cần đọc và kiểm tra toàn bộ đầu vào trước, rồi mới ghi cả bản ghi. Thao tác thêm cũng chèn dòng trước khi xử lý xong ngày hóa đơn, nên lỗi giữa chừng có thể để lại dòng dở dang.

### 4. Sinh ContractID có điểm gãy sau số 99

`modContract.bas:84` và `:91` chỉ nhận mẫu `-##`, sau đó lấy hai ký tự cuối. Khi đã tồn tại 01..99, lần tiếp theo sinh 100; lần sau bỏ qua 100 và lại sinh 100. Điều này áp dụng cả hợp đồng gốc theo năm/nhóm và con dưới cùng parent.

Cần đọc trọn hậu tố số và kiểm tra uniqueness ngay trước khi ghi.

### 5. GetSafeNumber loại bỏ cả dấu thập phân

`modContract.bas:161` và `:162` bỏ cả `.` và `,`. `frmPaymentEntry` dùng hàm này cho tiền và tỷ giá. Ví dụ chuỗi `1.25` trở thành `125`; `1,25` cũng thành `125`.

Nếu toàn bộ đầu vào được quy định là số nguyên thì cần khóa/giải thích quy tắc đó. Nếu cho phép tiền hoặc tỷ giá lẻ, phải có bộ phân tích số nhất quán với định dạng nhập; tránh tự suy đoán cả hai dấu là phân cách hàng nghìn.

### 6. Mở rộng mảng hai chiều bằng ReDim Preserve sai chiều

Ví dụ `frmReminder.frm:330` thay cận chiều thứ nhất của `buf(…, 0 To 9)` trong khi giữ dữ liệu. VBA chỉ cho thay cận chiều cuối với Preserve.

Mẫu tương tự có tại `frmContractSearch.frm:1567`, `:1590`, `:1696`, `:1710`; `frmPaymentSummary.frm:1536`; `modExportReport.bas:866`. Reminder kích hoạt nhánh này khi số bản ghi phù hợp vượt 1.000. Dữ liệu nhỏ có thể che lỗi.

### 7. Có lời gọi chưa tìm thấy định nghĩa trong 56 thành phần trích xuất

- `modContract`: QuickSortByDate, MoveContractFamily, GetDirectChildren, MoveLedgerBlock và hằng SIGNING_DATE_COL.
- `frmPaymentEntry.frm:125`: ExportSelectedLedgerVertical.
- `modPayment.bas:923`: hàm tên GetGroupFromContractArrold nhưng các phép gán kết quả dùng tên GetGroupFromContractArr.

Cần kiểm tra bằng Debug > Compile trên bản sao cùng đúng References/add-ins. Chưa khẳng định kết quả biên dịch thực tế vì không chạy VBE và chưa xác minh thư viện VBA bên ngoài. Một số đoạn có vẻ là code cũ chưa dọn.

### 8. Xóa hợp đồng chưa xử lý hồ sơ liên quan

`frmContractEdit.frm:617` chỉ xóa họ hợp đồng và ledger tương ứng. Không thấy bước xóa, chuyển hoặc chặn theo các bản ghi DOCUMENT_TRACKER trong handler này. Vì vậy có thể còn hồ sơ mang ContractID không còn tồn tại.

`ContractHasPayment` hiện kiểm tra giá trị J/L khác 0, không kiểm tra sự tồn tại LedgerID hoặc PaidAmount M. Với dữ liệu import hoặc dữ liệu đã sửa, điều kiện này chưa tương đương với “đã có bản ghi thanh toán”.

### 9. Parser AI phụ thuộc khoảng trắng JSON

`frmAskAI.frm:823` tìm đúng chuỗi `"content":"`; nhánh text tương tự. JSON hợp lệ dạng `"content": "…"` không khớp, có thể khiến phản hồi được trả về nhưng giao diện không lấy được nội dung. Đây là nhận xét từ parser, chưa gọi dịch vụ để kiểm thử.

API key và lịch sử key được ghi nguyên văn vào SETTINGS của AI_Data.xlsb (`frmAskAI.frm:146`, `:165`, `:198`); MaskKey chỉ che cách hiển thị. Cơ chế license cũng đặt bí mật và hàm sinh key trong client; cần thiết kế lại nếu mục tiêu là bảo vệ phân phối thương mại, không nên coi mã hóa XOR là ranh giới bảo mật.

### 10. Một số bất nhất khác cần kiểm tra tiếp

- `frmDocumentEdit.GetHoSoValue` trả `CON`, nhưng nhánh lưu ngày/trạng thái ở dòng 1849/1879 vẫn so với `HD`.
- Form ghi subtype 2 ở T, cache hồ sơ chỉ lấy A:S (`modGlobal.bas:64`). Cần kiểm tra các nơi tiêu thụ cache trước khi mở rộng schema.
- ExportPaymentSection tự thêm mọi B-CPB vào dictFilter ở dòng 1147, kể cả khi người dùng chọn một hợp đồng. Cần xác nhận đây có đúng yêu cầu báo cáo hay không.
- Nhánh fallback xuất thanh toán dùng tổng trên header và có thể tự gán ngày nếu không tìm được ngày hợp lệ; phải rà lại để tránh phân bổ sai kỳ hoặc quy đổi hai lần trên header đã là VND.
- `modBuildPackage.bas:374` xóa staging folder cả khi ZIP thất bại, nhưng thông báo dòng 395 nói folder chưa bị xóa.
- Nhiều hàm đặt lại Calculation/EnableEvents/DisplayAlerts về giá trị mặc định thay vì khôi phục trạng thái ban đầu. Một số handler nuốt lỗi bằng On Error Resume Next.
- QSM Link ghi toàn bộ mảng kết quả vào cột đích; dòng không khớp có phần tử rỗng. Đây là hành vi cần làm rõ trước khi dùng trên cột có nội dung hiện hữu.

## Hướng kiểm tra tiếp theo

1. Biên dịch bản sao với đúng References, xử lý định nghĩa thiếu và code cũ trước.
2. Thử khởi tạo trên thư mục dữ liệu sạch để xác nhận template/schema.
3. Kiểm tra vòng thêm/sửa thanh toán → chuyển dự án → mở lại, ngày sai, số thập phân, ID vượt 99, Reminder trên 1.000 hồ sơ.
4. Đối chiếu cùng một bộ dữ liệu giữa màn hình tổng hợp và file báo cáo, gồm phụ lục, Cancel, B-CPB, ngoại tệ và bộ lọc tháng/năm.

Không có sửa chữa nào được áp dụng vào workbook trong lượt đọc này.
