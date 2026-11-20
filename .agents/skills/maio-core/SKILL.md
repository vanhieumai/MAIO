---
name: maio-core
description: >-
  Hướng dẫn và quy chuẩn phát triển hệ thống quản lý web MAIO (chuyển đổi từ CPMS Plus, Quản lý Văn phòng, Chi tiêu).
  Sử dụng khi kiểm tra nghiệp vụ hợp đồng, dòng tiền thanh toán, cấu trúc dữ liệu, và kiến trúc web tương tự MLeague.
---

# MAIO Core Development Skill

Kỹ năng này hướng dẫn bất kỳ AI Assistant nào làm việc với dự án **MAIO** nắm bắt toàn bộ kiến thức nghiệp vụ, mô hình dữ liệu, phong cách thiết kế giao diện và quy trình mở rộng hệ thống.

---

## 1. Bản đồ Tham chiếu Dự án

Mọi tài liệu và mã nguồn nguồn gốc nằm trong dự án MAIO:

| Mục tiêu tra cứu | Đường dẫn tệp tin |
|---|---|
| Quy tắc phát triển | `PROJECT_RULES.md` |
| Kiến trúc hệ thống | `docs/ARCHITECTURE.md` |
| Lược đồ dữ liệu (Data Schema) | `docs/DATA_SCHEMA.md` |
| Lộ trình & Kế hoạch | `docs/MASTER_PLAN.md` |
| File Excel CPMS mới nhất | `legacy/cpms-plus/excel/CPMS AR - Sign Date Sort Fix.xlsm` |
| Dữ liệu CPMS thực tế | `legacy/cpms-plus/excel/CHC1_20261118_103621.cpms` |
| Mã nguồn VBA trích xuất | `legacy/cpms-plus/vba-source/` (56 files) |
| Đánh giá kỹ thuật CPMS | `legacy/cpms-plus/docs/CPMS_VBA_REVIEW.md` |
| Quản lý tài sản văn phòng | `legacy/office-finance/QuanLy_TaiSan_VanPhong_MoRong.xlsx` |
| Quản lý thu chi ngân hàng | `legacy/office-finance/QuanLy_Quy_ThuChi_NganHang.xlsx` |

---

## 2. Tiêu chuẩn Kỹ thuật & UI (Học hỏi từ MLeague)

1. **Stack:** HTML5 Semantic + CSS3 + Vanilla JavaScript (ES6 Modules) + IndexedDB + Firebase (Auth & Firestore).
2. **Typography & Styling:**
   - Font: Inter (Google Fonts).
   - Icons: Font Awesome 6 CDN (`https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css`).
   - Thiết kế dạng Component (Header, Sidebar Navigation, Card, Table, Modal, Drawer, Toast Notification).
3. **Phản hồi người dùng (User Feedback):**
   - Mọi thao tác Thêm / Sửa / Xóa dữ liệu phải có xác nhận (Confirmation Modal) và hiển thị thông báo kết quả (Toast) rõ ràng.
   - Luôn định dạng tiền tệ theo chuẩn Việt Nam Đồng (vd: `1.200.000.000 ₫`).
   - Định dạng ngày tháng: `DD/MM/YYYY`.

---

## 3. Quy chuẩn Nghiệp vụ Bắt buộc

### 3.1. Hợp đồng (Contracts)
- Phân nhóm: `Group A` (Khách hàng / Chủ đầu tư) và `Group B` (Nhà thầu phụ / NCC).
- Phụ lục (`Appendix`) kế thừa các thông tin của Hợp đồng cha và có `parent_id` liên kết.
- Cây hợp đồng hỗ trợ tính tổng giá trị của toàn bộ các phụ lục cộng dồn vào giá trị hợp đồng gốc.

### 3.2. Thanh toán (Payment Ledger)
- Công thức: `Chênh lệch còn lại = Đề nghị thanh toán - Thực thanh toán`.
- Bảng tổng hợp (Summary) gom nhóm theo hợp đồng cha và nhóm A/B.
- Hỗ trợ tỷ giá đối với các khoản thanh toán bằng ngoại tệ (USD, EUR,...).

### 3.3. Quản lý Văn phòng & Chi tiêu
- Tài sản phải có thông tin trạng thái: `Đang sử dụng`, `Trong kho`, `Bảo trì`, `Đã thanh lý`.
- Giao dịch thu chi phải thuộc về một Sổ quỹ / Tài khoản ngân hàng cụ thể và cập nhật tức thì số dư.
