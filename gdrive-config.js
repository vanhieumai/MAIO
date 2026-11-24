/**
 * Cấu hình kết nối Google Drive cho MAIO
 * Tương tự như firebase-config.js của MLeague.
 * Xem hướng dẫn thiết lập tại: docs/GOOGLE_DRIVE_SETUP.md
 */
window.MAIO_GDRIVE_CONFIG = {
  // Thay thế bằng Client ID từ Google Cloud Console (OAuth 2.0 Client ID for Web Application)
  clientId: "429965167454-7p2a5fodhue357s10chcussc8u45ufn0.apps.googleusercontent.com",
  
  // Thay thế bằng API Key từ Google Cloud Console
  apiKey: "AIzaSyAVqM4ADNcwXo-43eV-qf-oNaCHpD26qqo",
  
  // Tên thư mục gốc lưu trữ dữ liệu MAIO trên Google Drive của bạn
  appFolderName: "MAIO_Data",
  
  // Scope yêu cầu: drive.file chỉ cho phép MAIO đọc/ghi các file do chính MAIO tạo ra, đảm bảo an toàn tuyệt đối
  scopes: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/userinfo.email"
};
