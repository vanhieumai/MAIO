/**
 * Cấu hình kết nối Google Drive cho MAIO (File mẫu)
 * Sao chép file này thành `gdrive-config.js` và điền Client ID, API Key của bạn.
 * Xem hướng dẫn thiết lập tại: docs/GOOGLE_DRIVE_SETUP.md
 */
window.MAIO_GDRIVE_CONFIG = {
  // Thay thế bằng Client ID từ Google Cloud Console (OAuth 2.0 Client ID for Web Application)
  clientId: "YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com",
  
  // Thay thế bằng API Key từ Google Cloud Console
  apiKey: "YOUR_GOOGLE_API_KEY",
  
  // Tên thư mục gốc lưu trữ dữ liệu MAIO trên Google Drive của bạn
  appFolderName: "MAIO_Data",
  
  // Scope yêu cầu: drive.file chỉ cho phép MAIO đọc/ghi các file do chính MAIO tạo ra, đảm bảo an toàn tuyệt đối
  scopes: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email"
};
