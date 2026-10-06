"""
Bộ Chuyển Đổi CSV Sang Excel Chuẩn Ngân Hàng Để Gửi Khách Hàng (Client Exporter)
=================================================================================
Đảm bảo 100% bảo toàn số 0 đầu, định dạng tiền tệ #,##0, ngày tháng DD/MM/YYYY.
Sử dụng chung cấu hình định dạng từ app.exporter.excel_styles (DRY Principle).
"""

import os
import time
from pathlib import Path
from typing import Optional
import pandas as pd
import xlsxwriter

from app.core.config import DEFAULT_DATA_PATH, OUTPUT_CLIENT_EXCEL_PATH
from app.exporter.excel_styles import get_cctg_excel_formats, write_cctg_data_sheet


def export_csv_to_excel(
    csv_path: Optional[Path] = None,
    xlsx_path: Optional[Path] = None,
) -> bool:
    start_time = time.time()
    in_path = Path(csv_path) if csv_path else DEFAULT_DATA_PATH
    out_path = Path(xlsx_path) if xlsx_path else OUTPUT_CLIENT_EXCEL_PATH

    print("=" * 70, flush=True)
    print("TIẾN TRÌNH XUẤT BẢN EXCEL GỬI KHÁCH HÀNG (FULL CLIENT EXPORT)", flush=True)
    print("=" * 70, flush=True)

    if not in_path.exists():
        print(f"[LỖI] Không tìm thấy file đầu vào: {in_path}", flush=True)
        return False

    file_size_mb = os.path.getsize(in_path) / (1024 * 1024)
    print(f"[1/4] Đang nạp tệp dữ liệu: {in_path.name} ({file_size_mb:.2f} MB) ...", flush=True)

    if in_path.suffix.lower() == ".csv":
        df = pd.read_csv(in_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    else:
        df = pd.read_excel(in_path, dtype=str).fillna("")

    total_rows, total_cols = df.shape
    print(f"      -> Tổng số dòng: {total_rows:,} | Tổng số cột: {total_cols}", flush=True)

    print(f"\n[2/4] Đang khởi tạo Workbook Excel: {out_path.name} ...", flush=True)
    workbook = xlsxwriter.Workbook(out_path, {"constant_memory": True})
    worksheet = workbook.add_worksheet(in_path.stem)

    formats = get_cctg_excel_formats(workbook)

    print(f"\n[3/4] Đang ghi {total_rows:,} dòng dữ liệu kèm chuẩn hóa định dạng ...", flush=True)
    write_cctg_data_sheet(
        worksheet=worksheet,
        df=df,
        formats=formats,
        support_decimal_price=False,
        progress_callback=True,
    )

    print("\n[4/4] Đang lưu tệp Excel hoàn chỉnh ...", flush=True)
    workbook.close()

    elapsed = time.time() - start_time
    out_size_mb = os.path.getsize(out_path) / (1024 * 1024)

    print("=" * 70, flush=True)
    print(f"XUẤT BẢN THÀNH CÔNG: {out_path}", flush=True)
    print(f"Dung lượng: {out_size_mb:.2f} MB | Tổng số dòng: {total_rows:,} | Thời gian: {elapsed:.1f}s", flush=True)
    print("=" * 70, flush=True)
    return True


if __name__ == "__main__":
    export_csv_to_excel()
