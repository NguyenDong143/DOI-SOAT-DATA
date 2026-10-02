@echo off
chcp 65001 > nul
title Hệ Thống Đối Soát CCTG - API Server Backend
echo =====================================================================
echo   🚀 KHỞI CHẠY REST API SERVER ĐỐI SOÁT TAM GIÁC CCTG
echo =====================================================================
echo.
echo [1/2] Kiểm tra môi trường Python...
python --version
if %ERRORLEVEL% NEQ 0 (
    echo [LỖI] Không tìm thấy Python! Vui lòng cài đặt Python và thử lại.
    pause
    exit /b 1
)
echo.
echo [2/2] KHỞI ĐỘNG REST API SERVER (UVICORN)...
echo     - API Server:   http://localhost:8000
echo     - Swagger Docs: http://localhost:8000/docs
echo     - ReDoc:        http://localhost:8000/redoc
echo =====================================================================
echo Bấm Ctrl + C nếu muốn dừng Server.
echo.
python -m uvicorn app.api.main:app --host 0.0.0.0 --port 8000 --reload
pause
