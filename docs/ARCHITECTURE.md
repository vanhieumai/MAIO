# MAIO Architecture & Tech Stack (Kiến trúc Hệ thống MAIO)

Kiến trúc MAIO được phát triển dựa trên thành công của dự án **MLeague** kết hợp với nghiệp vụ chuyên sâu từ **CPMS Plus** và **Moffice**:

---

## 1. Triết lý Thiết kế Kế thừa từ MLeague

Dự án MLeague (`vanhieumai.github.io/Mleague`) có những ưu điểm kiến trúc xuất sắc:
1. **Zero-build / Low-friction (Không phụ thuộc bundler cồng kềnh):**
   - Viết bằng HTML5 + CSS3 hiện đại + Vanilla JavaScript ES6 Modules.
   - Không bị lỗi node_modules, không cần chạy `npm build` phức tạp mỗi khi kiểm thử, có thể chạy trực tiếp trên bất kỳ Web Server nào (Live Server, GitHub Pages, Vercel, Firebase Hosting, Nginx, Docker).
   - Dễ dàng cho bất kỳ AI nào (Gemini, Claude, GPT, Cursor) hay lập trình viên nào có thể đọc, hiểu, sửa đổi và kiểm tra trực tiếp mà không gặp lỗi môi trường/version package.
2. **Hybrid Data Layer (Local-first + Cloud BaaS):**
   - **Local-first**: Sử dụng `IndexedDB` hoặc `localStorage` lưu trữ tức thì, đảm bảo tốc độ phản hồi cực nhanh (<10ms), hoạt động cả khi offline, có thể Import/Export file JSON hoặc Excel bất kỳ lúc nào.
   - **Cloud Sync với Firebase**: Sử dụng **Firebase Authentication** (đăng nhập email/mật khẩu, phân quyền) + **Cloud Firestore** (đồng bộ dữ liệu đám mây đa thiết bị, chia sẻ link trực tiếp tương tự cơ chế chia sẻ lịch của MLeague).
3. **Giao diện người dùng Hiện đại & Đồng bộ (Design System):**
   - Font chữ chủ đạo: `Inter` (hiển thị tiếng Việt hoàn hảo, chuyên nghiệp cho ứng dụng quản trị).
   - Biểu tượng: `FontAwesome 6`.
   - Layout: App Shell với thanh điều hướng Header/Sidebar, Card container, Badge trạng thái rõ ràng, Modal tương tác nhanh, Data Table có tìm kiếm, lọc và phân trang mượt mà.

---

## 2. Cấu trúc Thư mục Dự án MAIO

```text
MAIO/
├── index.html                   # Điểm vào chính của Web App (App Shell)
├── css/
│   ├── main.css                 # Style toàn cục, biến màu sắc, typography (tương tự style.css của MLeague)
│   ├── components.css           # Thẻ Card, Modal, Table, Badge, Button, Tabs, Toast
│   └── modules/
│       ├── cpms.css             # Giao diện riêng cho Hợp đồng, Payment Ledger
│       ├── office.css           # Giao diện Quản lý Văn phòng & Tài sản
│       └── finance.css          # Giao diện Thu chi & Quỹ
├── js/
│   ├── app.js                   # Điểm khởi tạo ứng dụng, router và điều hướng phân hệ
│   ├── firebase-config.js       # Cấu hình kết nối Firebase (Auth + Firestore)
│   ├── core/
│   │   ├── db.js                # Quản lý Data Layer (IndexedDB local cache + Firestore sync)
│   │   ├── auth.js              # Xử lý đăng nhập, phiên người dùng, phân quyền
│   │   ├── event-bus.js         # Quản lý sự kiện giữa các module
│   │   └── utils.js             # Hàm tiện ích (format tiền tệ VND, định dạng ngày tháng, tính toán tỷ giá)
│   ├── modules/
│   │   ├── cpms/                # Phân hệ Quản lý Hợp đồng & Dự án (CPMS)
│   │   │   ├── contracts.js     # Logic cây hợp đồng A/B, phụ lục, tìm kiếm, sắp xếp
│   │   │   ├── payments.js      # Nhật ký thanh toán, nghiệm thu, tính chênh lệch, quy đổi
│   │   │   ├── documents.js     # Theo dõi công văn, sinh mã hồ sơ tự động
│   │   │   ├── reminders.js     # Quản lý cảnh báo và nhắc nhở hạn thanh toán/bảo lãnh
│   │   │   └── qsm.js           # Tiện ích liên kết số liệu QSM Toolkit
│   │   ├── office/              # Phân hệ Quản lý Văn phòng & Tài sản
│   │   │   ├── assets.js        # Quản lý danh mục tài sản, thiết bị
│   │   │   └── allocations.js   # Quản lý cấp phát, bảo trì, thu hồi
│   │   ├── finance/             # Phân hệ Chi tiêu & Quỹ
│   │   │   ├── cashbook.js      # Sổ quỹ tiền mặt và tài khoản ngân hàng
│   │   │   └── vouchers.js      # Quản lý phiếu thu, phiếu chi, tạm ứng
│   │   └── ai/
│   │       └── ask-ai.js        # Trợ lý AI hỏi đáp dữ liệu (kết nối Gemini API)
├── docs/                        # Toàn bộ tài liệu phân tích, kiến trúc, data schema
│   ├── MASTER_PLAN.md           # Kế hoạch tổng thể và lộ trình
│   ├── ARCHITECTURE.md          # Tài liệu kiến trúc này
│   ├── DATA_SCHEMA.md           # Đặc tả lược đồ dữ liệu chi tiết
│   └── FIREBASE_SETUP.md        # Hướng dẫn kết nối Firebase Authentication & Firestore
├── legacy/                      # Toàn bộ mã nguồn và file gốc để tham chiếu
│   ├── cpms-plus/               # Code VBA, Excel bản mới nhất, dữ liệu export thực tế
│   └── office-finance/          # File Excel quản lý tài sản văn phòng và thu chi
└── .agents/                     # Customizations dành cho Antigravity & các AI
    ├── skills/maio-core/        # Skill nghiệp vụ MAIO
    └── rules/                   # Quy tắc dự án
```

---

## 3. Luồng Dữ liệu & Cơ chế Hoạt động (Data Flow)

```
[ User Interaction ]
        │
        ▼
[ UI Components (HTML / CSS / JS) ]
        │
        ▼
[ Module Controllers (CPMS, Office, Finance, AI) ]
        │
        ▼
[ Core DB Adapter (js/core/db.js) ]
       /                             \
      ▼                               ▼
[ Local-First: IndexedDB ]      [ Cloud Sync: Firestore ]
  - Phản hồi tức thì (<10ms)      - Đồng bộ tài khoản
  - Hoạt động offline             - Phân quyền dự án
  - Import/Export Excel, JSON     - Chia sẻ link dữ liệu
```

---

## 4. Định hướng Triển khai (Deployment)
- Có thể deploy cực kỳ dễ dàng qua **GitHub Pages** (như MLeague: `vanhieumai.github.io/maio/`) hoặc **Firebase Hosting** / **Vercel** hoàn toàn miễn phí.
- Dữ liệu hoàn toàn an toàn: Cấu hình `firestore.rules` đảm bảo người dùng chỉ đọc/ghi dữ liệu của dự án mình được cấp quyền.
