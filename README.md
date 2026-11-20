# MAIO - All-in-One Management System (Hệ thống Quản lý Đa phân hệ)

Hệ thống quản lý tích hợp trên nền tảng Web chuyển đổi và nâng cấp từ **CPMS Plus (Add-in Excel)**, đồng thời mở rộng quản lý **Văn phòng**, **Chi tiêu / Quỹ** và **Hợp đồng - Dự án**.

---

## 1. Cấu trúc Tài nguyên Kế thừa (Legacy References)

Toàn bộ dữ liệu code và file mới nhất từ dự án CPMS Plus và các bảng tính quản lý văn phòng/chi tiêu đã được tập hợp trong thư mục `legacy/`:

- `legacy/cpms-plus/`
  - `excel/`:
    - `CPMS AR - Sign Date Sort Fix.xlsm` (Bản Excel/VBA mới nhất ngày 17/11/2026)
    - `CHC1_20261113_152509.cpms` (Dữ liệu dự án mẫu mới nhất ngày 13/11/2026)
  - `vba-source/`: Trích xuất 56 thành phần mã nguồn VBA (.bas, .cls, .frm) với hơn 38.000 dòng code.
  - `fixes/`: Các bản vá lỗi và cải tiến mới nhất (QSM Toolkit link, Expand Rows, Sign Date Sort, Ask AI...).
  - `docs/`:
    - `CPMS_VBA_REVIEW.md`: Đánh giá kiến trúc chi tiết, data model, quy tắc nghiệp vụ CPMS.
    - `CPMS_CHC1_DATA_REVIEW.md`: Đánh giá cấu trúc dữ liệu thực tế dự án CHC1.
- `legacy/office-finance/`
  - `QuanLy_Quy_ThuChi_NganHang.xlsx`: Quản lý thu chi tiền mặt, tài khoản ngân hàng, dòng tiền.
  - `QuanLy_TaiSan_VanPhong_MoRong.xlsx`: Quản lý tài sản trang thiết bị, công cụ dụng cụ, tài sản cố định.

---

## 2. Định hướng Kiến trúc Hệ thống Web MAIO

Hệ thống được thiết kế dạng Web Application hiện đại (Frontend SPA/SSR + Backend API RESTful/GraphQL + Cơ sở dữ liệu quan hệ), bao gồm 4 phân hệ chính:

### Phân hệ 1: Quản lý Hợp đồng & Dự án (CPMS Core)
Kế thừa trọn vẹn nghiệp vụ từ CPMS Plus:
- **Quản lý Hợp đồng**: Phân nhóm A (Chủ đầu tư/Khách hàng) và B (Nhà thầu phụ/Nhà cung cấp), cây hợp đồng phụ lục, trạng thái, giá trị hợp đồng, metadata.
- **Thanh toán & Giải ngân (Payment Ledger)**: Quản lý dòng tiền theo từng đợt nghiệm thu, đề nghị thanh toán, thực thanh toán, đối chiếu chênh lệch, quy đổi ngoại tệ.
- **Quản lý Hồ sơ & Công văn (Document Tracker)**: Tạo số công văn/hồ sơ theo quy tắc tự động, lưu trữ link tài liệu, kiểm soát trạng thái phát hành/nhận, phiên bản tài liệu.
- **Nhắc việc & Cảnh báo (Reminders)**: Hạn thanh toán, bảo lãnh hợp đồng, công nợ đến hạn, tiến độ hồ sơ.
- **Tiện ích QSM & Báo cáo**: Tổng hợp số liệu, xuất báo cáo động.

### Phân hệ 2: Quản lý Văn phòng & Tài sản (Office & Assets)
- Quản lý danh mục tài sản, thiết bị văn phòng, máy tính, máy in.
- Theo dõi cấp phát tài sản cho nhân sự/phòng ban.
- Quản lý lịch bảo trì, bảo dưỡng, khấu hao tài sản.
- Quản lý nhu cầu & cấp phát văn phòng phẩm.

### Phân hệ 3: Quản lý Chi tiêu & Quỹ (Finance & Expense)
- Quản lý quỹ tiền mặt (Cash Book) và các tài khoản ngân hàng (Bank Accounts).
- Quản lý phiếu thu, phiếu chi, tạm ứng, hoàn ứng.
- Phân loại danh mục chi phí (Chi phí văn phòng, chi phí dự án, chi phí vận hành).
- Báo cáo lưu chuyển tiền tệ, cân đối thu chi theo tháng/quý/năm.

### Phân hệ 4: Quản trị Hệ thống & Trợ lý AI (System & AI Assistant)
- Phân quyền người dùng theo vai trò (Admin, Kế toán, Quản lý dự án, Nhân viên văn phòng).
- Multi-tenancy / Nhiều dự án độc lập (tương tự chuyển đổi Project trong CPMS).
- **Trợ lý AI (MAIO AI Assistant)**: Tích hợp Gemini / OpenAI / Anthropic để tra cứu văn bản, phân tích hợp đồng và hỏi đáp dữ liệu tài chính dự án (kế thừa từ module Ask AI của CPMS Plus).
