# Đối chiếu data CHC1 với VBA CPMS

File đọc: `C:/Users/Personal/Desktop/CHC1_20260925_140917.cpms`.
Ngày đọc: 25/09/2026. Phân tích chỉ đọc, không chạy macro, không import vào CPMS và không sửa file nguồn.

## 1. Định dạng và quy mô

File `.cpms` này là workbook Excel Binary (XLSB) đổi phần mở rộng. Cấu trúc ZIP có `xl/workbook.bin`, worksheet `.bin`, sharedStrings và vbaProject. Điều này khớp với `frmHomePage.ExportSelectedProject`: lưu dữ liệu rồi FileCopy sang `.cpms`.

Có 9 sheet: CPMS, SYS_INFO, CONTRACT_DB, PAYMENT_LEDGER, DOCUMENT_TRACKER, SETTINGS, AI_NOTES, LIBRARY_DB, PROJECT_SETTINGS.

| Dữ liệu | Số lượng đã đọc |
| --- | ---: |
| Bản ghi hợp đồng/phụ lục/package | 63 |
| Hợp đồng MAIN | 1 |
| Hợp đồng SUB | 49 |
| Appendix, bao gồm package | 13 |
| Dòng tổng hợp ledger | 66 |
| Dòng giao dịch thanh toán | 196 |
| Hồ sơ | 153 |
| Loại tài liệu có cấu hình trong LIBRARY_DB | 13 |

66 dòng tổng hợp ledger gồm 63 dòng tương ứng CONTRACT_DB, 2 dòng GROUP-A/GROUP-B và 1 dòng chi phí nội bộ B-CPB-001. Trong 196 giao dịch có 19 dòng A, 166 dòng hợp đồng B và 11 dòng chi phí nội bộ.

Không phát hiện ID trùng trong ContractID, LedgerID của giao dịch và DocumentID. Mọi giao dịch đều tìm được header ledger tương ứng.

File data có 10 document modules VBA, chủ yếu là Worksheet_Deactivate đặt sheet thành veryHidden. Không có các standard module hoặc UserForm nghiệp vụ như file chương trình. Vì vậy data vẫn chứa VBA nhỏ phục vụ sheet, nhưng logic CPMS chính nằm trong file ứng dụng đã đọc trước đó.

## 2. Cấu hình dự án đã có trong data

PROJECT_SETTINGS xác định:

- ProjectID: CP260721174857.
- PROJ: CHC1.
- Dự án Nhà ở Chung cư Cao tầng Westlake Residential.
- LOCATION: Ô đất CHC1, Khu D6, Phú Thượng, Tây Hồ, Hà Nội.
- CLIENT=CLI, SUPCON=ART, MAINCON=EC.

`modLibrary.GetGlobalSetting` đọc các mã này từ PROJECT_SETTINGS. Mã nhà thầu riêng của từng hợp đồng nằm trong metadata cột Z của CONTRACT_DB, ví dụ `SUBCON=CTC`.

**Làm rõ nhận xét trước:** PROJECT_SETTINGS thiếu trong workbook chương trình CPMS AR.xlsm, nhưng có trong file data này. Do đó không được kết luận dự án CHC1 hiện thiếu cấu hình. Rủi ro khởi tạo template từ file chương trình trên máy sạch là một trường hợp khác.

SYS_INFO lưu DataVersion dạng `3.9`; PROJECT_SETTINGS lưu DataVersion là `1`. Đây là hai nơi ghi phiên bản khác nhau, chưa đủ căn cứ để kết luận dữ liệu lỗi hoặc phiên bản nào là chuẩn migration.

## 3. Cây hợp đồng là khóa để hiểu tổng hợp

Ví dụ thực tế:

```text
GROUP-A
└─ A-2026-01              hợp đồng MAIN, thứ tự A-1
   └─ A-2026-01-02        APP-06, thứ tự A-1.2
      ├─ A-2026-01-02-01  PKG-02, thứ tự A-1.2.1
      └─ A-2026-01-02-02  PKG-03, thứ tự A-1.2.2
```

CONTRACT_DB lưu ID nội bộ ở A, số hiển thị ở B, nhãn thứ tự ở X và thứ tự sắp xếp ở Y. Các giá trị Y của 63 bản ghi đều khác nhau. Thứ tự vật lý trên sheet không hoàn toàn trùng thứ tự giao diện: LoadContractList sort theo Y, còn RebuildContractSequence dựng thứ tự theo cây và ngày ký.

Giá trị APP-06 là 313.500.000.000; hai package lần lượt 156.231.479.624 và 157.268.520.376, cộng đúng bằng giá trị APP-06. Vì vậy cộng tất cả các dòng CONTRACT_DB sẽ đếm lặp trong trường hợp này. Không thể mặc định mọi giá trị phụ lục đều là giá trị tăng thêm ngoài hợp đồng gốc.

Năm trong ID `*-2026-*` cũng không phải năm nghiệp vụ thanh toán: PaymentDate của dữ liệu trải từ 01/12/2015 đến 03/09/2026. Bộ lọc tháng/năm phải dùng cột Q, không suy ra từ ID hoặc thời điểm tạo bản ghi.

## 4. Thanh toán: phân biệt dòng giao dịch và dòng tổng

Trong PAYMENT_LEDGER:

- B có LedgerID: dòng giao dịch. J:M là số tiền ở đơn vị của giao dịch, G là tỷ giá; W:Z là giá trị quy đổi.
- B trống, C có ContractID: dòng tổng hợp. A chỉ định cha. J:M lấy kết quả từ W:Z, đã là số tiền quy đổi.
- Header cộng các con trực tiếp có A bằng C của chính nó. Cơ chế này vừa cộng giao dịch trực tiếp vừa cộng subtotal của các phụ lục/package.

Đối chiếu ví dụ APP-06:

- PKG-02 đã trả: 114.091.854.196.
- PKG-03 đã trả: 95.656.031.667.
- APP-06 đã trả: 209.747.885.863, bằng tổng hai package.

Các ô tương ứng: PAYMENT_LEDGER M8, M15, M7. Root A-2026-01 còn có khoản trả trực tiếp ở dòng 4 và các phụ lục khác. Không được cộng cả header lẫn giao dịch khi tính tổng dự án.

### Ngoại tệ có tỷ giá theo từng lần trả

B-2026-14 có 3 giao dịch:

| Dòng ledger | Số gốc | Tỷ giá | Quy đổi |
| --- | ---: | ---: | ---: |
| 103 | 60.000 | 22.530 | 1.351.800.000 |
| 104 | 150.000 | 22.330 | 3.349.500.000 |
| 105 | 210.000 | 22.745 | 4.776.450.000 |
| Tổng header dòng 102 | | | 9.477.750.000 |

Như vậy phải quy đổi từng giao dịch rồi cộng. Không lấy tổng số gốc nhân duy nhất tỷ giá header, và không nhân lại tỷ giá vào M của header.

B-2026-18 còn có dòng bù số lẻ 7.900 VND với tỷ giá 1, nằm cùng hợp đồng có các lần trả ngoại tệ. Đây là trường hợp nghiệp vụ thực tế cần giữ khi sửa bộ xử lý tiền.

### Đối chiếu số đã lưu

Tự cộng các giao dịch sau quy đổi cho kết quả:

- Nhóm A: 463.612.767.074 VND.
- Nhóm B, gồm chi phí nội bộ: 450.609.369.607 VND.
- Trong B, B-CPB-001: 230.436.702 VND.

Các tổng này khớp giá trị đã lưu của GROUP-A và GROUP-B. Toàn bộ 196 dòng giao dịch có W:Z khớp J:M nhân G trong sai số kiểm tra 1 VND. Đây là đối chiếu cached values trong XLSB với phép tính độc lập, không phải chứng minh Excel đã tái tính công thức trong lượt này.

B-2026-07 có trạng thái Cancel. Giao dịch cũ vẫn giữ đề nghị thanh toán 108.155.599.656, nhưng header đưa W:Z về 0, phù hợp `UpdateLedgerFormulaByStatus`. Chênh lệch đó là chủ đích code, không nên tự sửa như lỗi cộng thiếu.

Dòng 78 có RequestAmount=0, PaidAmount=380.000.000, Remaining=-380.000.000. Điều này chứng minh dữ liệu hiện hữu không luôn tuân theo validation thêm mới yêu cầu RequestAmount>0; chưa kết luận đây là dữ liệu sai vì có thể là dữ liệu nhập trước hoặc điều chỉnh nghiệp vụ.

## 5. Hồ sơ và Reminder

153 hồ sơ gồm IPC 73, CON 49, APP 18, CMT 2, TRN 3, VO 3, OTHER 4 và CV 1.

Tình trạng: Accepted 73, Completed 29, Checking 6, Processing 3, Returned 4 và trống 38.

Với logic hiện tại của frmReminder và không lọc loại/từ khóa, tập hồ sơ chờ có 9 bản ghi Checking/Processing. Returned không thuộc bộ lọc chờ đó. Trong 9 bản ghi này có một bản ghi mang ContractID không còn có trong CONTRACT_DB.

Q/R đã có bên phát hành/nhận ở 111 bản ghi. S/T subtype đều trống trong bản data này, vì vậy vấn đề cache chỉ đọc tới S chưa thể hiện qua dữ liệu hiện tại.

Có 148 hồ sơ có Link. Phần lớn là đường dẫn ổ D của máy; export `.cpms` chỉ mang đường dẫn, không nhúng các tài liệu đính kèm. BuildDocumentZipPackage mới đọc đường dẫn để sao chép tài liệu thực tế vào gói ZIP. Chưa mở hoặc kiểm tra tồn tại các file liên kết.

### Bốn hồ sơ không nối được với CONTRACT_DB

| Dòng DOCUMENT_TRACKER | DocumentID | ContractID không tìm thấy |
| --- | --- | --- |
| 5 | DOC-20260603-004 | B-2026-29 |
| 141 | DOC-20260918-001 | A-2026-02 |
| 142 | DOC-20260918-002 | A-2026-02 |
| 148 | DOC-20260918-008 | A-2026-01-08 |

Hai ID A cũ xuất hiện trên hồ sơ có số/tên APP-17, trong khi APP-17 hiện có ContractID A-2026-01-09. Đây là dấu hiệu cần xác minh lịch sử đổi/xóa/tạo lại, chưa đủ để tự động gán lại.

`LoadDocumentList` nối bằng ContractID đúng giá trị. Vì vậy chọn APP-17 hiện tại sẽ không tự kéo các hồ sơ đang trỏ tới ID cũ. Điều này minh họa tác động của thiếu xử lý hồ sơ liên quan trong luồng xóa hợp đồng đã nêu ở bản review code.

## 6. Dữ liệu còn mang theo từ ứng dụng

SETTINGS trong file dự án vẫn chứa cấu hình AI và các trường API key/lịch sử key; không chép giá trị key vào báo cáo. AI_NOTES có một bản ghi ngoài tiêu đề. Trong code hiện tại, frmAskAI lại đọc dữ liệu qua modAISession từ AI_Data.xlsb riêng. Vì vậy không được mặc định các key/ghi chú trong project export này là cấu hình AI đang dùng ở phiên chạy hiện tại.

ExportSelectedProject sao chép cả workbook nên các sheet bổ trợ này cũng đi theo bản `.cpms`. Khi thiết kế export để chia sẻ dự án, cần quyết định rõ sheet nào thực sự thuộc dữ liệu dự án.

## Kết quả hiểu code sau khi đối chiếu

Dữ liệu xác nhận CPMS dùng workbook làm nơi lưu trữ, mảng/Dictionary làm cache, ContractID và ParentID làm quan hệ, còn UserForm điều phối đọc/ghi. Những quy tắc cần giữ khi sửa là: phân biệt ID và số hiển thị; phân biệt giao dịch và header; tổng hợp theo cây; quy đổi từng giao dịch; giữ lịch sử của hợp đồng Cancel; xử lý B-CPB riêng; và duy trì liên kết DOCUMENT_TRACKER khi thay đổi hợp đồng.

SHA-256 file nguồn: `899a2dcf54076b5ecbbc3faa85e7e0637fa7eb7f7e53875332e46b945fac6164`.
