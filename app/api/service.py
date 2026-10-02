"""
Service Layer for Reconciliation API
"""

import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
import pandas as pd

from app.core.config import OUTPUT_RECONCILE_PATH
from app.engine.reconcile_engine import TriangleReconciliationEngine


class ReconciliationService:
    _instance: Optional["ReconciliationService"] = None

    def __init__(self):
        self.engine: Optional[TriangleReconciliationEngine] = None
        self.last_run_timestamp: Optional[datetime] = None
        self.last_duration: float = 0.0
        self.is_running: bool = False

    @classmethod
    def get_instance(cls) -> "ReconciliationService":
        if cls._instance is None:
            cls._instance = ReconciliationService()
        return cls._instance

    def run_reconciliation(
        self,
        data_path: Optional[str] = None,
        hd_path: Optional[str] = None,
        sk_path: Optional[str] = None,
        output_path: Optional[str] = None,
    ) -> Dict[str, Any]:
        if self.is_running:
            raise RuntimeError("Tiến trình đối soát đang chạy, vui lòng đợi hoàn tất.")

        self.is_running = True
        start_time = time.time()
        try:
            engine = TriangleReconciliationEngine(
                data_path=Path(data_path) if data_path else None,
                hd_path=Path(hd_path) if hd_path else None,
                sk_path=Path(sk_path) if sk_path else None,
                output_path=Path(output_path) if output_path else None,
            )
            engine.execute()

            self.engine = engine
            self.last_run_timestamp = datetime.now()
            self.last_duration = round(time.time() - start_time, 2)

            return self.get_summary()
        finally:
            self.is_running = False

    def get_summary(self) -> Dict[str, Any]:
        if self.engine is None:
            self.run_reconciliation()

        m = self.engine.summary_metrics
        tot = m["total_rows"]
        return {
            "execution_time": self.last_run_timestamp.strftime("%Y-%m-%d %H:%M:%S") if self.last_run_timestamp else "",
            "duration_seconds": self.last_duration,
            "total_rows": tot,
            "count_golden": m["count_golden"],
            "pct_golden": round(m["count_golden"] / tot * 100, 2),
            "count_future_sk": m["count_future_sk"],
            "count_abbank": m["count_abbank"],
            "count_pending_sale": m["count_pending_sale"],
            "count_discrepancy": m["count_discrepancy"],
            "count_missing_hd": m["count_missing_hd"],
            "sk_max_date": m["sk_max_date"],
            "sk_dupes_removed": m.get("sk_dupes_removed", 0),
            "data_file_name": m["data_file_name"],
            "count_rounding": m.get("count_rounding", 0),
            "count_retail_rounding_all": m.get("count_retail_rounding_all", 0),
            "count_retail_rounding_t9": m.get("count_retail_rounding_t9", 0),
            "sum_retail_data_t9": m.get("sum_retail_data_t9", 0),
            "sum_retail_bank_t9": m.get("sum_retail_bank_t9", 0),
            "sum_retail_diff_t9": m.get("sum_retail_diff_t9", 0),
            "output_file": str(self.engine.output_path),
        }

    def get_retail_rounding_records_paginated(
        self,
        page: int = 1,
        page_size: int = 50,
        period: str = "all",
        search: Optional[str] = None,
    ) -> Dict[str, Any]:
        if self.engine is None:
            self.run_reconciliation()

        records = self.engine.retail_rounding_contracts if period == "t9" else self.engine.retail_rounding_all
        if search:
            s = search.lower()
            records = [
                r for r in records
                if any(s in str(v).lower() for v in r.values())
            ]

        total_items = len(records)
        total_pages = max(1, (total_items + page_size - 1) // page_size)
        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size

        return {
            "total_items": total_items,
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages,
            "items": records[start_idx:end_idx],
        }

    def get_problem_records_paginated(
        self,
        page: int = 1,
        page_size: int = 50,
        search: Optional[str] = None,
        level: str = "contract",
    ) -> Dict[str, Any]:
        if self.engine is None:
            self.run_reconciliation()

        records = self.engine.problem_contracts if level == "contract" else self.engine.problem_rows
        if search:
            s = search.lower()
            records = [
                r for r in records
                if any(s in str(v).lower() for v in r.values())
            ]

        total_items = len(records)
        total_pages = max(1, (total_items + page_size - 1) // page_size)
        start_idx = (page - 1) * page_size
        end_idx = start_idx + page_size

        return {
            "total_items": total_items,
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages,
            "items": records[start_idx:end_idx],
        }

    def get_export_file_path(self) -> Path:
        if self.engine and self.engine.output_path.exists():
            return self.engine.output_path
        return OUTPUT_RECONCILE_PATH
