"""
Định Dạng & Kiểu Dáng Bảng Tính Excel Chuẩn Ngân Hàng (Shared Excel Styles)
========================================================================
Tập trung toàn bộ cấu hình màu sắc, phông chữ, định dạng số, độ rộng cột
dùng chung cho Client Exporter và Clean Core Data Pipeline (DRY Principle).
"""

from typing import Any, Dict, List, Set
import pandas as pd
import xlsxwriter

from app.core.parsers import clean_int_str, parse_date_str

# Nhóm các cột phân loại theo kiểu dữ liệu
MONEY_COLS: Set[str] = {
    "Mệnh giá",
    "Tổng GT HĐ mua",
    "Tổng GT HĐ Bán",
    "Đơn giá bán/ số seri",
    "Lãi Coupon đã trả thực tế",
}

INT_COLS: Set[str] = {
    "Số lượng",
    "Thời hạn nắm giữ (tháng)",
}

RATE_COLS: Set[str] = {
    "LS Coupon",
    "Lãi suất HĐ Mua",
}

DATE_COLS: Set[str] = {
    "Ngày báo cáo",
    "Ngày phát hành AZ",
    "Ngày đáo hạn AZ",
    "Ngày mua (ngày nắm giữ)",
    "Ngày hết hạn nắm giữ",
    "Ngày bán",
}

CENTER_TEXT_COLS: Set[str] = {
    "Số CIF",
    "Loại KH",
    "Số định danh/MST",
    "Mã CCTG",
    "Số sổ AZ",
    "Số FT HĐ Mua",
    "Số FT HĐ Bán",
    "Số hoá đơn",
    "Ký hiệu hoá đơn",
    "Trạng thái hoá đơn",
    "Mã CKS HĐ Mua",
    "Trạng thái GD Mua",
    "Mã CKS HĐ Bán",
    "Trạng thái GD Bán",
}


def get_cctg_excel_formats(workbook: xlsxwriter.Workbook, font_family: str = "Arial") -> Dict[str, Any]:
    """Tạo bộ định dạng ô tiêu chuẩn ngân hàng cho Workbook XlsxWriter."""
    return {
        "header": workbook.add_format({
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
        }),
        "left": workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "left",
            "valign": "vcenter",
        }),
        "center": workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "center",
            "valign": "vcenter",
        }),
        "currency": workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0",
        }),
        "currency_decimal": workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0.00",
        }),
        "integer": workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0",
        }),
        "rate": workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "right",
            "valign": "vcenter",
            "num_format": "0.0",
        }),
        "date": workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "center",
            "valign": "vcenter",
        }),
    }


def setup_cctg_worksheet_columns(worksheet: Any, headers: List[str]) -> None:
    """Thiết lập độ rộng cột tối ưu cho bảng dữ liệu CCTG."""
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
        elif h in DATE_COLS:
            w = 15
        worksheet.set_column(col_idx, col_idx, w)


def write_cctg_data_sheet(
    worksheet: Any,
    df: pd.DataFrame,
    formats: Dict[str, Any],
    support_decimal_price: bool = True,
    progress_callback: bool = False,
) -> None:
    """
    Ghi dữ liệu DataFrame chuẩn 34 cột vào Worksheet với đầy đủ định dạng số,
    ngày tháng và bảo toàn số 0 đứng đầu cho các mã định danh.
    """
    headers = list(df.columns)
    total_rows = len(df)
    total_cols = len(headers)

    # 1. Định cấu hình cột & Freeze Panes
    worksheet.hide_gridlines(0)
    worksheet.freeze_panes(1, 4)
    setup_cctg_worksheet_columns(worksheet, headers)

    # 2. Ghi Header
    worksheet.set_row(0, 28)
    for col_idx, h in enumerate(headers):
        worksheet.write(0, col_idx, h, formats["header"])

    # 3. Ghi dữ liệu từng dòng
    data_matrix = df.values
    for r_idx in range(total_rows):
        if progress_callback and r_idx % 25_000 == 0 and r_idx > 0:
            print(f"      ... đã hoàn thành {r_idx:,}/{total_rows:,} dòng ({r_idx/total_rows*100:.1f}%)", flush=True)

        row_num = r_idx + 1
        worksheet.set_row(row_num, 19)
        row_raw = data_matrix[r_idx]

        for col_idx, val in enumerate(row_raw):
            h_name = headers[col_idx]

            if not val or str(val).lower() in ("nan", "none", "null"):
                fmt = formats["center"] if h_name in CENTER_TEXT_COLS else formats["left"]
                worksheet.write_string(row_num, col_idx, "", fmt)
                continue

            val_str = str(val).strip()

            if support_decimal_price and h_name == "Đơn giá bán/ số seri":
                try:
                    num_v = float(val_str.replace(",", ""))
                    if num_v.is_integer():
                        worksheet.write_number(row_num, col_idx, int(num_v), formats["currency"])
                    else:
                        worksheet.write_number(row_num, col_idx, num_v, formats["currency_decimal"])
                except ValueError:
                    worksheet.write_string(row_num, col_idx, val_str, formats["left"])

            elif h_name in MONEY_COLS:
                try:
                    num_v = float(val_str.replace(",", ""))
                    worksheet.write_number(row_num, col_idx, num_v, formats["currency"])
                except ValueError:
                    worksheet.write_string(row_num, col_idx, val_str, formats["left"])

            elif h_name in INT_COLS:
                try:
                    num_v = int(float(val_str.replace(",", "")))
                    worksheet.write_number(row_num, col_idx, num_v, formats["integer"])
                except ValueError:
                    worksheet.write_string(row_num, col_idx, val_str, formats["left"])

            elif h_name in RATE_COLS:
                try:
                    num_v = float(val_str)
                    worksheet.write_number(row_num, col_idx, num_v, formats["rate"])
                except ValueError:
                    worksheet.write_string(row_num, col_idx, val_str, formats["center"])

            elif h_name in DATE_COLS:
                worksheet.write_string(row_num, col_idx, parse_date_str(val_str), formats["date"])

            elif h_name in ("Số hoá đơn", "Mã CKS HĐ Mua", "Mã CKS HĐ Bán"):
                worksheet.write_string(row_num, col_idx, clean_int_str(val_str), formats["center"])

            elif h_name in CENTER_TEXT_COLS:
                worksheet.write_string(row_num, col_idx, val_str, formats["center"])

            else:
                worksheet.write_string(row_num, col_idx, val_str, formats["left"])

    worksheet.autofilter(0, 0, total_rows, total_cols - 1)
