@echo off
chcp 65001 >nul
echo ============================================================
echo   MAIO - CẬP NHẬT ADD-IN CPMS PLUS (CPMS AR.xlam)
echo ============================================================
echo.
echo Đang kiểm tra tiến trình Excel...
tasklist /FI "IMAGENAME eq excel.exe" 2>NUL | find /I /N "excel.exe">NUL
if "%ERRORLEVEL%"=="0" (
    echo [CẢNH BÁO] Microsoft Excel đang chạy!
    echo Vui lòng lưu công việc và ĐÓNG Excel lại, sau đó bấm phím bất kỳ để tiếp tục.
    pause >nul
)

set "SRC=%~dp0excel\CPMS AR.xlam"
set "DEST=%APPDATA%\Microsoft\AddIns\CPMS AR.xlam"

if not exist "%SRC%" (
    echo [LỖI] Không tìm thấy tệp nguồn: "%SRC%"
    pause
    exit /b 1
)

copy /Y "%SRC%" "%DEST%" >nul
if "%ERRORLEVEL%"=="0" (
    echo.
    echo [THÀNH CÔNG] Đã cập nhật CPMS AR.xlam vào thư mục Add-ins thành công!
    echo Đường dẫn: %DEST%
    echo.
    echo Bạn có thể mở Excel và sử dụng tính năng Import CPMS bình thường.
) else (
    echo.
    echo [LỖI] Không thể ghi đè file. Hãy chắc chắn rằng Excel đã được đóng hoàn toàn.
)
echo.
pause
