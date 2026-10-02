@echo off
chcp 65001 > nul
title Đối Soát Tam Giác CCTG - Hóa Đơn - Sao Kê TK43
echo =====================================================================
echo   📊 CHẠY ĐỐI SOÁT TAM GIÁC (DATA CCTG ^<--^> HÓA ĐƠN ^<--^> SAO KÊ TK43)
echo =====================================================================
echo.
echo Đang xử lý đối soát đa chiều trên hệ thống module app.engine ...
python -m app.engine.reconcile_engine
if %ERRORLEVEL% EQU 0 (
    echo.
    echo [THÀNH CÔNG] Báo cáo 5 sheet đã được xuất tại:
    echo              data\output\Data_abba_Chuan_Va_DoiSoat.xlsx
    echo.
    set /p OPEN_NOW="Bạn có muốn mở file Báo Cáo Đối Soát ngay bây giờ không? (Y/N): "
    if /i "%OPEN_NOW%"=="Y" (
        start "" "data\output\Data_abba_Chuan_Va_DoiSoat.xlsx"
    )
) else (
    echo.
    echo [LỖI] Quá trình đối soát gặp sự cố!
)
echo.
pause
