@echo off
chcp 65001 > nul
title Xuất Dữ Liệu Sạch Import CD CORE DATA
echo =====================================================================
echo   🚀 TIẾN TRÌNH LÀM SẠCH DỮ LIỆU ĐỂ IMPORT VÀO CD CORE DATA
echo =====================================================================
echo.
echo Đang nạp dữ liệu và thực thi pipeline làm sạch chuyên sâu ...
echo 1. Khớp và nạp hóa đơn điện tử còn thiếu từ BC Hoa don.xlsx
echo 2. Bổ sung FT Mua ngân hàng từ saoke_TK43.xlsx
echo 3. Sửa dứt điểm xung đột HĐ 18715 Lâm Gia Phước (CIF 13558581)
echo 4. Phân bổ chênh lệch làm tròn 264 HĐ bán lẻ vào Đơn giá bán
echo.
python -m app.engine.clean_core_pipeline
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [THÀNH CÔNG] Dữ liệu sạch đã được xuất bản tại:
    echo   1. File CSV Core:   data\output\Data_abba_Clean_Core.csv
    echo   2. File Excel Core: data\output\Data_abba_Clean_Core.xlsx
    echo.
    set /p OPEN_NOW="Bạn có muốn mở file Excel Core ngay bây giờ không? (Y/N): "
    if /i "%OPEN_NOW%"=="Y" (
        start "" "data\output\Data_abba_Clean_Core.xlsx"
    )
) else (
    echo.
    echo [LỖI] Quá trình làm sạch dữ liệu gặp sự cố!
)
echo.
pause
