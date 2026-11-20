# Hướng dẫn Thiết lập Google Drive Cloud Sync cho MAIO

Tài liệu này hướng dẫn bạn cách tạo **Client ID** và **API Key** miễn phí từ Google Cloud để kết nối Google Drive làm kho lưu trữ đám mây cho MAIO.

---

## 1. Tổng quan & Ưu điểm

* **Bảo mật tuyệt đối:** Ứng dụng chỉ sử dụng quyền `https://www.googleapis.com/auth/drive.file` — nghĩa là MAIO **chỉ có thể đọc và ghi các file do chính MAIO tạo ra**, hoàn toàn KHÔNG THỂ xem hoặc can thiệp vào các tài liệu, ảnh, video khác có sẵn trên Google Drive của bạn.
* **Tự động tạo thư mục:** Khi kết nối thành công lần đầu, MAIO sẽ tự động tạo thư mục `MAIO_Data/` trên Google Drive của bạn để lưu trữ các file dữ liệu dự án.
* **Chi phí:** Hoàn toàn **0 đồng** (sử dụng gói miễn phí của Google Cloud và 15GB miễn phí của tài khoản Google cá nhân).

---

## 2. Các bước thiết lập trên Google Cloud Console

### Bước 1: Tạo dự án mới trên Google Cloud
1. Truy cập: [Google Cloud Console](https://console.cloud.google.com/).
2. Đăng nhập bằng tài khoản Google cá nhân của bạn.
3. Ở thanh tiêu đề trên cùng, bấm vào danh sách chọn Project -> Bấm **New Project** (Dự án mới).
4. Đặt tên dự án: ví dụ `MAIO-System` rồi bấm **Create**.

### Bước 2: Bật Google Drive API
1. Trong menu bên trái (biểu tượng ☰), chọn **APIs & Services** → **Library** (Thư viện).
2. Tìm kiếm từ khóa: `Google Drive API`.
3. Bấm vào **Google Drive API** và chọn **Enable** (Bật).

### Bước 3: Cấu hình Màn hình đồng ý OAuth (OAuth consent screen)
1. Trong menu bên trái, chọn **APIs & Services** → **OAuth consent screen**.
2. Chọn loại đối tượng người dùng: **External** (Bên ngoài) rồi bấm **Create**.
3. Điền các thông tin cơ bản:
   * **App name:** `MAIO Management System`
   * **User support email:** Chọn email của bạn.
   * **Developer contact email:** Điền email của bạn.
4. Bấm **Save and Continue** qua các bước Scopes (có thể để mặc định).
5. Ở bước **Test users** (Người dùng thử nghiệm): Bấm **Add Users** và thêm chính email Google của bạn vào (để được phép đăng nhập khi app đang ở trạng thái Testing).
6. Bấm **Save and Continue** để hoàn tất.

### Bước 4: Tạo OAuth 2.0 Client ID (Dành cho Web)
1. Trong menu bên trái, chọn **APIs & Services** → **Credentials** (Thông tin xác thực).
2. Bấm **+ Create Credentials** (Tạo thông tin xác thực) → Chọn **OAuth client ID**.
3. Mục **Application type** (Loại ứng dụng): Chọn **Web application**.
4. Mục **Name**: Đặt tên `MAIO Web Client`.
5. Mục **Authorized JavaScript origins** (Nguồn gốc JavaScript được ủy quyền): Thêm các địa chỉ sau:
   * `http://localhost`
   * `http://localhost:5500` (nếu dùng VS Code Live Server)
   * `http://localhost:8000` (nếu dùng Python HTTP server)
   * `http://127.0.0.1:5500`
   * `https://vanhieumai.github.io` (Địa chỉ GitHub Pages của bạn khi deploy online)
6. Bấm **Create**. Bạn sẽ nhận được một chuỗi **Client ID** (dạng `xxxxxxxxxxxx.apps.googleusercontent.com`).

### Bước 5: Tạo API Key (Tùy chọn nhưng khuyến nghị)
1. Tại trang **Credentials**, bấm tiếp **+ Create Credentials** → Chọn **API key**.
2. Copy chuỗi API Key vừa được tạo.

---

## 3. Cập nhật cấu hình vào dự án MAIO

Mở file [gdrive-config.js](../gdrive-config.js) và điền Client ID cùng API Key của bạn:

```javascript
window.MAIO_GDRIVE_CONFIG = {
  clientId: "DIEN_CLIENT_ID_CUA_BAN_TAI_DAY.apps.googleusercontent.com",
  apiKey: "DIEN_API_KEY_CUA_BAN_TAI_DAY",
  appFolderName: "MAIO_Data",
  scopes: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email"
};
```

Sau khi điền, bạn chỉ cần mở `index.html` trên trình duyệt và bấm **Kết nối Google Drive** để hoàn tất!

---

## 4. Cấu trúc Cây Thư Mục Chuẩn Hóa trên Google Drive

MAIO tự động thiết lập và quản trị cây thư mục theo mô hình 3 trụ cột (Quản lý Dự án CPMS, Văn phòng Moffice, Sổ quỹ Chi tiêu) với cấu trúc phân nhóm chuyên sâu:

```text
📁 MAIO_Data/
├── 📁 01_QuanLy_DuAn/
│   └── 📁 DuAn_{MaDA}/                          # Thư mục riêng cho từng Dự án (VD: DuAn_CHC1)
│       ├── 📁 01_Nhom_A_ChuDauTu/               # Hợp đồng với Chủ đầu tư / Tổng thầu
│       │   └── 📁 {SoHieuHD}/                   # Thư mục cho từng hợp đồng A
│       │       ├── 📁 HopDong_PhuLuc/           # 📌 Thư mục "Hợp đồng, Phụ lục"
│       │       │   ├── 📁 HopDong_Goc/          # Chứa các file scan hợp đồng gốc, biên bản đàm phán (nhiều file)
│       │       │   ├── 📁 PhuLuc_01/            # Thư mục riêng cho Phụ lục số 01 (file scan & tài liệu)
│       │       │   └── 📁 PhuLuc_02/            # Thư mục riêng cho Phụ lục số 02
│       │       ├── 📁 ThanhToan/                # 📌 Thư mục "Thanh toán" (Các đợt IPC, nghiệm thu, thanh toán)
│       │       └── 📁 Doc/                      # 📌 Thư mục "Doc" (Lưu các tài liệu khác: Công văn, Bảo lãnh, Pháp lý)
│       │
│       └── 📁 02_Nhom_B_ThauPhu_NCC/            # Hợp đồng Nhà thầu phụ & Nhà cung cấp
│           │                                    # Phân loại thành 8 nhóm chuyên biệt:
│           ├── 📁 Construction/                 # Thi công xây lắp (cọc, hầm, kết cấu thân, hoàn thiện)
│           ├── 📁 Procurement/                  # Mua sắm vật tư (thép, xi măng, nhôm kính...)
│           ├── 📁 Consultancy/                  # Tư vấn, giám sát, thẩm tra, kiểm định
│           ├── 📁 Design/                       # Thiết kế kiến trúc, kết cấu, MEP
│           ├── 📁 Equipment/                    # Thiết bị cơ điện, thang máy, trạm điện
│           ├── 📁 Service/                      # Dịch vụ kỹ thuật, trắc đạc, bảo vệ, vệ sinh
│           ├── 📁 Maintenance/                  # Bảo dưỡng, bảo trì kỹ thuật định kỳ
│           └── 📁 Rental/                       # Thuê máy móc, cẩu tháp, giàn giáo, cốp pha
│               │
│               └── 📁 {SoHieuHD}_{TenDoiTac}/   # Thư mục cụ thể của hợp đồng B
│                   ├── 📁 HopDong_PhuLuc/       # 📌 Thư mục "Hợp đồng, Phụ lục"
│                   │   ├── 📁 HopDong_Goc/      # Chứa các file scan hợp đồng chính (nhiều file scan)
│                   │   ├── 📁 PhuLuc_01/        # Thư mục riêng phụ lục 01 (file scan & dự toán phụ lục)
│                   │   └── 📁 PhuLuc_02/        # Thư mục riêng phụ lục 02
│                   ├── 📁 ThanhToan/            # 📌 Thư mục "Thanh toán" (Hồ sơ thanh quyết toán đợt)
│                   └── 📁 Doc/                  # 📌 Thư mục "Doc" (Lưu tài liệu khác: Công văn, Bảo lãnh, Kỹ thuật)
│
├── 📁 02_QuanLy_VanPhong/                       # Quản lý văn phòng & Hành chính
│   ├── 📁 01_TaiSan_TrangThietBi/               # Danh mục tài sản, biên bản bàn giao thiết bị
│   ├── 📁 02_HoSo_HanhChinh_CapPhat/            # Biên bản cấp phát, luân chuyển văn phòng phẩm
│   └── 📁 03_HopDong_Thue_VanPhong/             # Hợp đồng thuê mặt bằng văn phòng, dịch vụ
│
└── 📁 03_QuanLy_ChiTieu/                        # Quản lý dòng tiền & Thu chi
    ├── 📁 01_PhieuChi_ChungTu/                  # Phiếu chi, hóa đơn đỏ VAT, chứng từ chuyển khoản
    ├── 📁 02_PhieuThu/                          # Phiếu thu tiền mặt, chứng từ báo có ngân hàng
    └── 📁 03_BaoCao_SoQuy/                      # Báo cáo quỹ, sao kê ngân hàng hàng tháng
```
