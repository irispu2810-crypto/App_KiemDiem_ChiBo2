@echo off
chcp 65001 > nul
title Chia Se Mang LAN & Ma QR - Chi Bo 2 VietinBank Bao Loc
echo ======================================================================
echo   CHI BỘ 2 - ĐẢNG BỘ NGÂN HÀNG TMCP CÔNG THƯƠNG VN CN BẢO LỘC
echo        CỔNG CHIA SẺ MÃ QR & KẾT NỐI MẠNG NỘI BỘ (LAN / WI-FI)
echo ======================================================================
echo.
echo Đang dò tìm IP LAN và khởi động máy chủ chia sẻ...

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo Khởi chạy máy chủ chia sẻ và mở Cổng kết nối QR Code...
    python "%~dp0tao_qr_chibo2.py"
    start "" "http://localhost:8080/share_portal.html"
    python "%~dp0server.py"
) else (
    echo Khởi chạy trực tiếp Cổng chia sẻ trên trình duyệt...
    start "" "%~dp0share_portal.html"
)
pause
