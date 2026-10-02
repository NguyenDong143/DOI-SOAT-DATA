@echo off
chcp 65001 > nul
title Chuyển Đổi File CSV CCTG Sang Excel Chuẩn Ngân Hàng
echo =====================================================================
echo   📊 CHUYỂN ĐỔI CSV DATA CCTG SANG EXCEL (.XLSX) ĐỂ GỬI KHÁCH HÀNG
echo =====================================================================
echo.
echo Đang xử lý chuyển đổi dữ liệu từ data\input sang data\output ...
python -m app.exporter.excel_exporter
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [THÀNH CÔNG] File Excel đã sẵn sàng để gửi khách hàng tại:
    echo              data\output\Data_abba_20261002.xlsx
    echo.
    set /p OPEN_NOW="Bạn có muốn mở file Excel ngay bây giờ không? (Y/N): "
    if /i "%OPEN_NOW%"=="Y" (
        start "" "data\output\Data_abba_20261002.xlsx"
    )
) else (
    echo.
    echo [LỖI] Có lỗi trong quá trình chuyển đổi.
)
echo.
pause
