@echo off
chcp 65001 >nul
title MAIO Local Server
echo ========================================================
echo  MAIO - HỆ THỐNG QUẢN LÝ ĐA PHÂN HỆ
echo  Đang khởi động máy chủ cục bộ: http://127.0.0.1:5500
echo ========================================================
start http://127.0.0.1:5500
python server.py 5500
if %errorlevel% neq 0 (
    echo.
    echo Không tìm thấy Python hoặc lỗi khởi động. Đang chuyển sang http.server...
    python -m http.server 5500 --bind 127.0.0.1
)
pause
