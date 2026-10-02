"""
Bộ Chuyển Đổi CSV Sang Excel Chuẩn Ngân Hàng Để Gửi Khách Hàng (Client Exporter)
=================================================================================
Đảm bảo 100% bảo toàn số 0 đầu, định dạng tiền tệ #,##0, ngày tháng DD/MM/YYYY.
"""

import os
import time
from pathlib import Path
from typing import Optional
import pandas as pd
import xlsxwriter

from app.core.config import DEFAULT_DATA_PATH, OUTPUT_CLIENT_EXCEL_PATH
from app.core.parsers import clean_int_str, parse_date_str


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

    worksheet.hide_gridlines(0)
    worksheet.freeze_panes(1, 4)

    font_family = "Arial"
    fmt_header = workbook.add_format({
        "bold": True,
        "bg_color": "#1B365D",
        "font_color": "#FFFFFF",
        "font_name": font_family,
        "font_size": 10,
        "border": 1,
        "border_color": "#D9D9D9",
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
    })

    fmt_text_left = workbook.add_format({
        "font_name": font_family,
        "font_size": 9,
        "border": 1,
        "border_color": "#E0E0E0",
        "align": "left",
        "valign": "vcenter",
    })

    fmt_text_center = workbook.add_format({
        "font_name": font_family,
        "font_size": 9,
        "border": 1,
        "border_color": "#E0E0E0",
        "align": "center",
        "valign": "vcenter",
    })

    fmt_currency = workbook.add_format({
        "font_name": font_family,
        "font_size": 9,
        "border": 1,
        "border_color": "#E0E0E0",
        "align": "right",
        "valign": "vcenter",
        "num_format": "#,##0",
    })

    fmt_integer = workbook.add_format({
        "font_name": font_family,
        "font_size": 9,
        "border": 1,
        "border_color": "#E0E0E0",
        "align": "right",
        "valign": "vcenter",
        "num_format": "#,##0",
    })

    fmt_percent = workbook.add_format({
        "font_name": font_family,
        "font_size": 9,
        "border": 1,
        "border_color": "#E0E0E0",
        "align": "right",
        "valign": "vcenter",
        "num_format": "0.0",
    })

    fmt_date = workbook.add_format({
        "font_name": font_family,
        "font_size": 9,
        "border": 1,
        "border_color": "#E0E0E0",
        "align": "center",
        "valign": "vcenter",
    })

    headers = list(df.columns)
    money_cols = {"Mệnh giá", "Tổng GT HĐ mua", "Đơn giá bán/ số seri", "Tổng GT HĐ Bán", "Lãi Coupon đã trả thực tế"}
    int_cols = {"Số lượng", "Thời hạn nắm giữ (tháng)"}
    rate_cols = {"LS Coupon", "Lãi suất HĐ Mua"}
    date_cols = {"Ngày báo cáo", "Ngày phát hành AZ", "Ngày đáo hạn AZ", "Ngày mua (ngày nắm giữ)", "Ngày hết hạn nắm giữ", "Ngày bán"}
    center_text_cols = {
        "Số CIF", "Loại KH", "Số định danh/MST", "Mã CCTG", "Số sổ AZ", "Số FT HĐ Mua", "Số FT HĐ Bán",
        "Số hoá đơn", "Ký hiệu hoá đơn", "Trạng thái hoá đơn", "Mã CKS HĐ Mua", "Trạng thái GD Mua",
        "Mã CKS HĐ Bán", "Trạng thái GD Bán"
    }

    # Độ rộng cột tối ưu
    for col_idx, h in enumerate(headers):
        w = max(len(h) + 4, 12)
        if h in ("Số series thứ cấp", "Mã tra cứu fkey"):
            w = 38
        elif h in ("Số HĐ mua", "Số HĐ Bán"):
            w = 28
        elif h in ("Tên KH",):
            w = 25
        elif h in ("Tổng GT HĐ mua", "Tổng GT HĐ Bán", "Đơn giá bán/ số seri"):
            w = 18
        elif h in date_cols:
            w = 15
        worksheet.set_column(col_idx, col_idx, w)

    worksheet.set_row(0, 28)
    for col_idx, h in enumerate(headers):
        worksheet.write(0, col_idx, h, fmt_header)

    print(f"\n[3/4] Đang ghi {total_rows:,} dòng dữ liệu kèm chuẩn hóa định dạng ...", flush=True)
    data_matrix = df.values
    for r_idx in range(total_rows):
        if r_idx % 25_000 == 0 and r_idx > 0:
            print(f"      ... đã hoàn thành {r_idx:,}/{total_rows:,} dòng ({r_idx/total_rows*100:.1f}%)", flush=True)

        row_num = r_idx + 1
        worksheet.set_row(row_num, 19)
        row_raw = data_matrix[r_idx]

        for col_idx, val in enumerate(row_raw):
            h_name = headers[col_idx]

            if not val or str(val).lower() in ("nan", "none", "null"):
                worksheet.write_string(row_num, col_idx, "", fmt_text_center if h_name in center_text_cols else fmt_text_left)
                continue

            if h_name in money_cols:
                try:
                    num_val = float(str(val).replace(",", "").strip())
                    worksheet.write_number(row_num, col_idx, num_val, fmt_currency)
                except ValueError:
                    worksheet.write_string(row_num, col_idx, str(val), fmt_text_left)

            elif h_name in int_cols:
                try:
                    int_val = int(float(str(val).strip()))
                    worksheet.write_number(row_num, col_idx, int_val, fmt_integer)
                except ValueError:
                    worksheet.write_string(row_num, col_idx, str(val), fmt_text_left)

            elif h_name in rate_cols:
                try:
                    r_val = float(str(val).strip())
                    worksheet.write_number(row_num, col_idx, r_val, fmt_percent)
                except ValueError:
                    worksheet.write_string(row_num, col_idx, str(val), fmt_text_center)

            elif h_name in date_cols:
                worksheet.write_string(row_num, col_idx, parse_date_str(str(val)), fmt_date)

            elif h_name in ("Số hoá đơn", "Mã CKS HĐ Mua", "Mã CKS HĐ Bán"):
                worksheet.write_string(row_num, col_idx, clean_int_str(val), fmt_text_center)

            elif h_name in center_text_cols:
                worksheet.write_string(row_num, col_idx, str(val), fmt_text_center)

            else:
                worksheet.write_string(row_num, col_idx, str(val), fmt_text_left)

    worksheet.autofilter(0, 0, total_rows, total_cols - 1)

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
