"""
Pydantic Schemas for Reconciliation REST API
"""

from typing import Any, Dict, List, Optional
from pydantic import BaseModel, Field


class HealthResponse(BaseModel):
    status: str = "ok"
    version: str = "2.0.0"
    service: str = "Triangle Reconciliation & Data Processing API"


class ReconcileRunRequest(BaseModel):
    data_path: Optional[str] = Field(default=None, description="Đường dẫn file dữ liệu CCTG (.csv/.xlsx)")
    hd_path: Optional[str] = Field(default=None, description="Đường dẫn file Hóa đơn (.xlsx)")
    sk_path: Optional[str] = Field(default=None, description="Đường dẫn file Sao kê TK43 (.xlsx)")
    output_path: Optional[str] = Field(default=None, description="Đường dẫn file Excel đầu ra (.xlsx)")


class ReconcileSummaryResponse(BaseModel):
    execution_time: str
    duration_seconds: float
    total_rows: int
    count_golden: int
    pct_golden: float
    count_future_sk: int
    count_abbank: int
    count_pending_sale: int
    count_discrepancy: int
    count_missing_hd: int
    sk_max_date: str
    sk_dupes_removed: int = 0
    count_rounding: int = 0
    count_retail_rounding_all: int = 0
    count_retail_rounding_t9: int = 0
    sum_retail_data_t9: float = 0.0
    sum_retail_bank_t9: float = 0.0
    sum_retail_diff_t9: float = 0.0
    data_file_name: str
    output_file: str


class PaginatedRecordsResponse(BaseModel):
    total_items: int
    page: int
    page_size: int
    total_pages: int
    items: List[Dict[str, Any]]
