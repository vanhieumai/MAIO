# MAIO Data Schema & Domain Models (Đặc tả Dữ liệu Chi tiết)

Tài liệu này chuẩn hóa toàn bộ cấu trúc dữ liệu từ các hệ thống gốc:
1. **CPMS Plus** (Add-in Excel: `CONTRACT_DB`, `PAYMENT_LEDGER`, `DOCUMENT_TRACKER`, `PROJECT_SETTINGS`)
2. **Moffice - Văn phòng** (`QuanLy_TaiSan_VanPhong_MoRong.xlsx`)
3. **Moffice - Thu chi** (`QuanLy_Quy_ThuChi_NganHang.xlsx`)

---

## 1. Phân hệ CPMS: Hợp đồng & Dự án (CPMS Core)

### 1.1. Bảng Dự án (`projects`)
| Trường (Field) | Kiểu dữ liệu | Mô tả / Diễn giải | Nguồn gốc CPMS |
|---|---|---|---|
| `id` | String (UUID / Code) | Mã định danh dự án (vd: `CHC1`) | `ProjectID` |
| `name` | String | Tên dự án đầy đủ | `ProjectName` |
| `code` | String | Ký hiệu viết tắt (vd: `CHC1`) | `ProjectCode` |
| `location` | String | Địa điểm xây dựng | `Location` |
| `client` | String | Chủ đầu tư | `ClientName` |
| `description` | String | Mô tả tóm tắt dự án | `Description` |
| `created_at` | Timestamp | Ngày tạo | `CreatedDate` |
| `updated_at` | Timestamp | Ngày cập nhật | System |

---

### 1.2. Bảng Hợp đồng & Phụ lục (`contracts`)
Đại diện cho cấu trúc cây hợp đồng: Main Contract, Appendix, Package con.

| Trường (Field) | Kiểu dữ liệu | Mô tả / Quy ước nghiệp vụ | Cột CPMS (`CONTRACT_DB`) |
|---|---|---|---|
| `id` | String | Khóa chính nội bộ (vd: `A-2026-01`, `A-2026-01-01`, `B-2026-01`) | Cột 1 (`ContractID`) |
| `project_id` | String | Khóa ngoại trỏ đến `projects.id` | Multi-project |
| `group` | Enum: `A`, `B` | `A`: Chủ đầu tư / Tổng thầu; `B`: Thầu phụ / Nhà cung cấp | Cột 1 (`GROUP-A` / `GROUP-B`) |
| `parent_id` | String (nullable) | Trỏ đến Contract cha (nếu là Phụ lục hoặc Gói thầu con) | Suy ra từ cấu trúc ContractID |
| `contract_no` | String | Số hợp đồng / Số phụ lục thực tế (vd: `68/2022/HĐKT/CHC1/VEC`, `APP-05`) | Cột 2 (`ContractNo.`) |
| `contract_type` | Enum / String | `MAIN` (HĐ chính), `Appendix` (Phụ lục), `Package` (Gói thầu), `Internal` | Cột 3 (`ContractType`) |
| `signing_date` | Date | Ngày ký hợp đồng | Cột 4 (`SigningDate`) |
| `end_date` | Date (nullable) | Ngày kết thúc hiệu lực | Cột 5 (`EndDate`) |
| `status` | Enum | `Active`, `Draft`, `Closed`, `Liquidated`, `Terminated` | Cột 6 (`Status`) |
| `company_a` | String | Tên pháp nhân Bên A | Cột 7 (`Company1`) |
| `company_b` | String | Tên pháp nhân Bên B | Cột 8 (`Company2`) |
| `scope_of_work`| Text | Nội dung phạm vi công việc gói thầu | Cột 11 (`ScopeOfWork`) |
| `currency` | String | Loại tiền tệ (`VND`, `USD`,...) mặc định `VND` | Metadata |
| `exchange_rate`| Number | Tỷ giá quy đổi (mặc định 1.0) | Cột 13 (`ExchangeRate`) |
| `estimated_val`| Number | Giá trị dự toán / phê duyệt | Cột 14 (`EstimatedValue`) |
| `contract_val` | Number | Giá trị hợp đồng ký kết | Cột 15 (`ContractValue`) |
| `rep_a` | String | Người đại diện bên A | Cột 16 (`RepA`) |
| `rep_b` | String | Người đại diện bên B | Cột 17 (`RepB`) |
| `category` | String | Nhóm chuyên môn (vd: `Design & Build`, `Construction`, `Procurement`)| Cột 18 (`Categories`) |
| `payee` | String | Đơn vị thụ hưởng thanh toán | Cột 19 (`Payee`) |
| `bank_account` | String | Số tài khoản thụ hưởng | Cột 20 (`BankAccountNo.`) |
| `bank_name` | String | Tên ngân hàng & chi nhánh | Cột 21 (`BankName`) |
| `contact_b` | String | Người liên hệ phụ trách | Cột 22 (`ContactPersonB`) |
| `display_no` | String | Mã hiển thị phân cấp (vd: `A-1`, `A-1.1`, `A-1.2`) | Cột 24 |
| `sort_order` | Number | Thứ tự sắp xếp hiển thị | Cột 25 |

---

### 1.3. Bảng Nhật ký Thanh toán (`payment_ledgers`)
Ghi nhận từng đợt tạm ứng, nghiệm thu và giải ngân thanh toán thực tế.

| Trường (Field) | Kiểu dữ liệu | Mô tả / Quy ước nghiệp vụ | Cột CPMS (`PAYMENT_LEDGER`) |
|---|---|---|---|
| `id` | String | Mã dòng nhật ký (vd: `LG-20260123162820`) | Cột 2 (`LedgerID`) |
| `contract_id` | String | Khóa ngoại trỏ đến `contracts.id` | Cột 3 (`ContractID`) |
| `parent_id` | String (nullable) | ID dòng cha hoặc HĐ cha tổng hợp | Cột 1 (`RowID` / `ParentID`) |
| `item_index` | Number | Đợt thanh toán (vd: Đợt 1, 2, 3...) | Cột 5 (`ItemIndex`) |
| `item_name` | String | Diễn giải nội dung đợt (vd: Tạm ứng đợt 1, Nghiệm thu đợt 2) | Cột 6 (`ItemName`) |
| `exchange_rate`| Number | Tỷ giá tại thời điểm thanh toán | Cột 7 (`ExchangeRate`) |
| `contract_val` | Number | Giá trị hợp đồng tại thời điểm ghi nhận | Cột 9 (`ContractValue`) |
| `actual_volume`| Number | Giá trị khối lượng thực hiện lũy kế | Cột 10 (`ActualVolume`) |
| `accepted_vol` | Number | Giá trị khối lượng đã nghiệm thu A-B | Cột 11 (`AcceptedVolume`) |
| `request_amount`| Number | Số tiền bên B đề nghị thanh toán | Cột 12 (`RequestAmount`) |
| `paid_amount` | Number | Số tiền bên A thực tế đã giải ngân | Cột 13 (`PaidAmount`) |
| `remaining` | Number | Chênh lệch còn lại (`request_amount` - `paid_amount`) | Cột 14 (`Remaining`) |
| `payment_ref` | String | Số chứng từ thanh toán (Ủy nhiệm chi / Giấy báo nợ) | Cột 15 (`PaymentRef`) |
| `payment_no` | String | Đợt/Kỳ thanh toán (vd: `01`, `02`) | Cột 16 (`PaymentNo.`) |
| `payment_date`| Date (nullable) | Ngày thanh toán | Cột 17 (`PaymentDate`) |
| `invoice_value`| Number | Giá trị hóa đơn VAT xuất | Cột 18 (`InvoiceValue`) |
| `invoice_ref` | String | Ký hiệu & số hóa đơn GTGT | Cột 19 (`InvoiceRef`) |
| `invoice_date`| Date (nullable) | Ngày xuất hóa đơn | Cột 20 (`InvoiceDate`) |
| `note` | Text | Ghi chú chi tiết đợt thanh toán | Cột 21 (`Note`) |

---

### 1.4. Bảng Theo dõi Hồ sơ & Công văn (`documents`)
Quản lý công văn đến/đi, hồ sơ nghiệm thu, thanh toán, pháp lý.

| Trường (Field) | Kiểu dữ liệu | Mô tả / Quy ước nghiệp vụ | Cột CPMS (`DOCUMENT_TRACKER`) |
|---|---|---|---|
| `id` | String | Mã hồ sơ tự động (vd: `DOC-2026-001`) | Cột 1 (`DocumentID`) |
| `contract_id` | String (nullable) | Hợp đồng liên quan (nếu có) | Cột 2 (`ContractID`) |
| `category` | String | Phân loại (Nghiệm thu, Thanh toán, Pháp lý, Thiết kế...) | Cột 3 (`Hồ sơ`) |
| `doc_no` | String | Số hiệu công văn/tài liệu chính thức | Cột 4 (`Số`) |
| `title` | String | Tiêu đề / Tên hồ sơ | Cột 5 (`Tên hồ sơ`) |
| `file_link` | String | Đường dẫn tệp tin lưu trữ (Drive / Cloud / Local) | Cột 6 (`Link`) |
| `start_date` | Date | Ngày phát hành / Ngày có hiệu lực | Cột 7 (`Từ ngày`) |
| `due_date` | Date (nullable) | Hạn hoàn thành / Ngày hết hạn | Cột 8 (`Đến ngày`) |
| `received_date`| Date (nullable) | Ngày tiếp nhận hồ sơ | Cột 9 (`Tiếp nhận`) |
| `return_date` | Date (nullable) | Ngày trả kết quả duyệt | Cột 10 (`Trả lại`) |
| `recipient` | String | Người/Đơn vị tiếp nhận | Cột 11 (`Người nhận lại`) |
| `quantity` | Number | Số lượng bộ tài liệu bàn giao | Cột 12 (`Số lượng`) |
| `attachments` | Text | Danh sách tài liệu đính kèm | Cột 13 (`Hồ sơ kèm theo`) |
| `iteration` | Number | Lần nộp / Phiên bản đệ trình (Lần 1, 2...) | Cột 14 (`Lần`) |
| `status` | Enum | `Draft`, `Submitted`, `UnderReview`, `Approved`, `Rejected` | Cột 15 (`Tình trạng`) |
| `notes` | Text | Ghi chú tiến trình xử lý | Cột 16 (`Ghi chú`) |

---

## 2. Phân hệ Quản lý Văn phòng & Tài sản (Office & Assets)

Dựa trên dữ liệu từ `QuanLy_TaiSan_VanPhong_MoRong.xlsx`:

### 2.1. Bảng Tài sản Thiết bị (`office_assets`)
- `id`: Mã tài sản (vd: `TS-VP-001`, `TS-IT-002`)
- `asset_name`: Tên tài sản (vd: Laptop Dell XPS 15, Bàn làm việc chữ L)
- `category`: Nhóm tài sản (Thiết bị IT, Nội thất văn phòng, Máy văn phòng, Phương tiện)
- `serial_number`: Số seri máy / Mã vạch
- `purchase_date`: Ngày mua
- `original_price`: Giá trị mua ban đầu
- `warranty_months`: Thời hạn bảo hành (tháng)
- `status`: Tình trạng (`DangSuDung`, `TrongKho`, `DangBaoTri`, `DaThanhLy`)
- `assigned_to`: Nhân sự đang giữ / bàn giao
- `department`: Phòng ban quản lý / sử dụng
- `location`: Vị trí đặt (Phòng họp, Tầng 3, Kho...)
- `notes`: Ghi chú kỹ thuật

### 2.2. Bảng Lịch sử Cấp phát & Bảo trì (`asset_logs`)
- `id`: Mã log
- `asset_id`: Trỏ đến `office_assets.id`
- `action_type`: `CapPhat`, `ThuHoi`, `BaoTri`, `SuaChua`, `ThanhLy`
- `action_date`: Ngày thực hiện
- `performed_by`: Người thực hiện / Kỹ thuật viên
- `receiver`: Người nhận bàn giao
- `cost`: Chi phí sửa chữa/bảo dưỡng (nếu có)
- `notes`: Nội dung biên bản bàn giao hoặc tình trạng hư hỏng

---

## 3. Phân hệ Quản lý Chi tiêu & Sổ quỹ (Finance & Expenses)

Dựa trên dữ liệu từ `QuanLy_Quy_ThuChi_NganHang.xlsx`:

### 3.1. Bảng Tài khoản / Quỹ (`financial_accounts`)
- `id`: Mã tài khoản (vd: `ACC-CASH`, `ACC-VCB-01`, `ACC-TCB-02`)
- `account_name`: Tên tài khoản / Quỹ (vd: Quỹ tiền mặt văn phòng, VCB TK Công ty)
- `account_type`: `Cash` (Tiền mặt), `Bank` (Ngân hàng), `CreditCard`
- `bank_name`: Tên ngân hàng & Chi nhánh
- `account_number`: Số tài khoản ngân hàng
- `opening_balance`: Số dư đầu kỳ
- `current_balance`: Số dư hiện tại
- `currency`: Mặc định `VND`

### 3.2. Bảng Giao dịch Thu chi (`cash_transactions`)
- `id`: Mã giao dịch (vd: `TX-202611-001`)
- `account_id`: Trỏ đến `financial_accounts.id`
- `tx_date`: Ngày phát sinh giao dịch
- `tx_type`: `Thu` (Phiếu thu), `Chi` (Phiếu chi), `TamUng` (Tạm ứng), `HoanUng`, `ChuyenKhoanNoiBo`
- `amount`: Số tiền giao dịch
- `category`: Khoản mục thu chi (Chi phí thuê VP, Điện nước, Tiếp khách, Mua VPP, Trả tiền thầu,...)
- `related_project_id`: Liên kết đến Dự án CPMS (nếu khoản chi cho dự án cụ thể)
- `payer_receiver`: Đối tượng nộp tiền / Người thụ hưởng nhận tiền
- `created_by`: Người lập phiếu
- `approved_by`: Người phê duyệt
- `status`: `Draft`, `PendingApproval`, `Approved`, `Cancelled`
- `voucher_code`: Số phiếu thu/chi nội bộ (vd: `PT-11/01`, `PC-11/05`)
- `notes`: Diễn giải chi tiết lý do chi
- `attachments`: Link chứng từ, hóa đơn đính kèm
