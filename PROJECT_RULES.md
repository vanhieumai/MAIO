# QUY TẮC PHÁT TRIỂN DỰ ÁN MAIO (PROJECT RULES FOR ALL AI AGENTS)

Tài liệu này là KIM CHỈ NAM bắt buộc mọi AI (Gemini, Claude, GPT, Cursor, Copilot, Antigravity) và lập trình viên phải tuân thủ nghiêm ngặt khi phát triển dự án **MAIO**.

---

## 1. Giao tiếp & Ngôn ngữ
- **Luôn giao tiếp với người dùng bằng Tiếng Việt**.
- Mọi ghi chú code, tài liệu và giao diện người dùng phải hiển thị Tiếng Việt rõ ràng, chuẩn mực và chuyên nghiệp.

---

## 2. Triết lý Kiến trúc Công nghệ (Kế thừa từ MLeague)
- **Zero-build / Low-friction Web SPA:**
  - Sử dụng chuẩn Web hiện đại: **HTML5 + CSS3 + Vanilla JavaScript (ES6+ Modules)**.
  - **TUYỆT ĐỐI KHÔNG** cài đặt các framework có quy trình build quá phức tạp làm phát sinh hàng trăm MB `node_modules` trừ khi người dùng có yêu cầu rõ ràng.
  - Mục tiêu: Web app phải chạy được ngay lập tức trên bất kỳ môi trường nào (Live Server, GitHub Pages, Firebase Hosting) mà không bị phụ thuộc phiên bản thư viện hay hỏng bản quyền.
- **Hybrid Data Layer (Local-first + Cloud BaaS):**
  - **Local-first**: Sử dụng `IndexedDB` để lưu trữ dữ liệu cục bộ với tốc độ phản hồi tức thì, hỗ trợ offline 100%, có thể Import/Export JSON và Excel dễ dàng.
  - **Cloud Sync**: Tích hợp với **Firebase Authentication & Cloud Firestore** (như MLeague) để đồng bộ dữ liệu đám mây và chia sẻ báo cáo qua link.
- **Giao diện (UI/UX) Đồng bộ:**
  - Font: `Inter`
  - Icons: `FontAwesome 6`
  - Style: Thẻ card bo góc, bảng dữ liệu hiện đại, modal tương tác nhanh, toast notification thông báo kết quả.

---

## 3. Quy tắc Nghiệp vụ Cốt lõi (Bảo tồn từ CPMS Plus)
- **Hợp đồng (Contracts):**
  - Luôn duy trì phân nhóm: **Nhóm A** (Chủ đầu tư / Khách hàng / Tổng thầu) và **Nhóm B** (Thầu phụ / Nhà cung cấp).
  - Phụ lục hợp đồng (Appendix) phải liên kết đúng vào Hợp đồng cha (`parent_id`).
  - Hỗ trợ đầy đủ các trường: Ngày ký, ngày hết hạn, giá trị hợp đồng, tỷ giá, người đại diện hai bên, tài khoản ngân hàng thụ hưởng.
- **Sổ Thanh toán (Payment Ledger):**
  - Tự động tính toán chuẩn xác: `Chênh lệch còn lại = Đề nghị thanh toán - Thực tế đã thanh toán`.
  - Giá trị thanh toán quy đổi = Giá trị * Tỷ giá.
  - Không được làm tròn số tiền sai lệch gây sai số kế toán.
- **Hồ sơ & Công văn (Document Tracker):**
  - Tuân thủ cấu trúc sinh mã hồ sơ tự động từ CPMS.
  - Phân loại trạng thái hồ sơ: Dự thảo, Đã nộp, Đang duyệt, Đã duyệt, Bị từ chối.
- **Quản lý Văn phòng & Chi tiêu (Office & Finance):**
  - Tuân thủ lược đồ tại `docs/DATA_SCHEMA.md`.

---

## 4. Quy trình Khi Sửa Đổi Code
1. **Kiểm tra trước khi thay đổi:** Đọc kỹ tài liệu trong `docs/` và mã nguồn trong `legacy/` trước khi sửa hoặc tạo mới.
2. **Không tự ý xóa dữ liệu cũ:** Dữ liệu gốc trong `legacy/` là tài sản tham chiếu bất biến, tuyệt đối không được xóa hoặc ghi đè làm mất dấu vết.
3. **Báo cáo trung thực:** Luôn kiểm tra kết quả thực thi và giải thích rõ ràng những gì đã làm cho người dùng.
