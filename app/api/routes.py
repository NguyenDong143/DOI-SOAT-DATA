"""
API Endpoints Router for Reconciliation Portal
"""

from typing import Optional
from fastapi import APIRouter, HTTPException, Query
from fastapi.responses import FileResponse

from app.api.schemas import (
    HealthResponse,
    PaginatedRecordsResponse,
    ReconcileRunRequest,
    ReconcileSummaryResponse,
)
from app.api.service import ReconciliationService

router = APIRouter(prefix="/api", tags=["Đối Soát Tam Giác CCTG"])
service = ReconciliationService.get_instance()


@router.get("/health", response_model=HealthResponse)
async def health_check():
    """Kiểm tra tình trạng hoạt động của API Service."""
    return HealthResponse()


@router.post("/reconcile/run", response_model=ReconcileSummaryResponse)
async def run_reconciliation(req: Optional[ReconcileRunRequest] = None):
    """Kích hoạt chạy tiến trình đối soát tam giác."""
    try:
        data_p = req.data_path if req else None
        hd_p = req.hd_path if req else None
        sk_p = req.sk_path if req else None
        out_p = req.output_path if req else None
        return service.run_reconciliation(data_p, hd_p, sk_p, out_p)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/reconcile/summary", response_model=ReconcileSummaryResponse)
async def get_summary():
    """Lấy dữ liệu tổng quan tỷ lệ đạt và thống kê kết quả đối soát."""
    try:
        return service.get_summary()
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/reconcile/problems", response_model=PaginatedRecordsResponse)
async def get_problem_records(
    page: int = Query(default=1, ge=1, description="Số trang"),
    page_size: int = Query(default=50, ge=1, le=500, description="Số dòng mỗi trang"),
    search: Optional[str] = Query(default=None, description="Tìm kiếm từ khóa (FT, HĐ, Tên KH, Sổ AZ)"),
    level: str = Query(default="contract", description="Cấp độ: 'contract' (Tổng hợp theo Hợp đồng, không trùng) hoặc 'passbook' (Chi tiết từng sổ AZ)"),
):
    """Lấy danh sách các dòng có vấn đề / sai lệch (có phân trang và tìm kiếm)."""
    try:
        return service.get_problem_records_paginated(page=page, page_size=page_size, search=search, level=level)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/reconcile/retail-rounding", response_model=PaginatedRecordsResponse)
async def get_retail_rounding_records(
    page: int = Query(default=1, ge=1, description="Số trang"),
    page_size: int = Query(default=50, ge=1, le=500, description="Số dòng mỗi trang"),
    period: str = Query(default="all", description="Kỳ đối soát: 'all' (Toàn bộ 264 HĐ làm tròn) hoặc 't9' (Tháng 9/2026 - 230 HĐ)"),
    search: Optional[str] = Query(default=None, description="Tìm kiếm mã HĐ hoặc Tên KH"),
):
    """Lấy danh sách các hợp đồng bán lẻ có chênh lệch làm tròn số học (đối chiếu 1:1 với file chuẩn)."""
    try:
        return service.get_retail_rounding_records_paginated(page=page, page_size=page_size, period=period, search=search)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/reconcile/export")
async def export_excel():
    """Tải xuống tệp Excel báo cáo đối soát 6 sheet hoàn chỉnh (.xlsx)."""
    try:
        file_path = service.get_export_file_path()
        if not file_path.exists():
            raise HTTPException(status_code=404, detail="File báo cáo chưa được tạo.")
        return FileResponse(
            path=str(file_path),
            filename=file_path.name,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

