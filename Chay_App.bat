@echo off
chcp 65001 > nul
title Chi Bo 2 - VietinBank Bao Loc - App Kiem Diem Dang Vien
echo ======================================================================
echo   CHI BỘ 2 - ĐẢNG BỘ NGÂN HÀNG TMCP CÔNG THƯƠNG VN CN BẢO LỘC
echo      HỆ THỐNG KIỂM ĐIỂM, ĐÁNH GIÁ, XẾP LOẠI ĐẢNG VIÊN VÀ TẬP THỂ
echo ======================================================================
echo.
echo Đang khởi chạy ứng dụng...

where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo Khởi động máy chủ Python hỗ trợ kết nối mạng LAN và xuất Word...
    python "%~dp0server.py"
) else (
    echo Khởi chạy trực tiếp trên trình duyệt Web mặc định...
    start "" "%~dp0app_kiem_diem_chi_bo_2.html"
)
pause
