"""
Tiện Ích Chuẩn Hóa Dữ Liệu & Bóc Tách (Data Parsers & Normalizers)
================================================================
Tập trung toàn bộ logic DRY dùng chung cho Engine, Exporter và API.
"""

import re
from datetime import date, datetime
from typing import Any, List, Optional, Tuple
import pandas as pd

from app.core.config import RE_FT


def norm_code(val: Any) -> str:
    """Loại bỏ ký tự đặc biệt, dấu nối và đưa về chữ in hoa."""
    if pd.isna(val) or val is None:
        return ""
    return re.sub(r"[^A-Z0-9]", "", str(val).upper())


def extract_ft(val: Any) -> str:
    """Bóc tách mã FT chuẩn dạng FT..."""
    if pd.isna(val) or val is None:
        return ""
    m = RE_FT.search(str(val))
    return m.group(1).upper() if m else ""


def parse_date(v: Any) -> Optional[date]:
    """Chuẩn hóa đa dạng các kiểu dữ liệu ngày về datetime.date an toàn (loại trừ NaT)."""
    if pd.isna(v) or v is None or v is pd.NaT:
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        try:
            return v.date()
        except Exception:
            return None
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none", "nat", ""):
        return None
    if len(s) == 8 and s.isdigit():
        try:
            return datetime.strptime(s, "%Y%m%d").date()
        except Exception:
            pass
    for fmt in (
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y-%m-%d",
        "%m/%d/%Y %I:%M:%S %p",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except Exception:
            pass
    try:
        dt = pd.to_datetime(s, dayfirst=True)
        if pd.isna(dt) or dt is pd.NaT:
            return None
        return dt.date()
    except Exception:
        pass
    return None


def dates_match(dt1: Optional[date], dt2: Optional[date]) -> Tuple[bool, int]:
    """
    Kiểm tra khớp ngày có tính dung sai chuyển tiếp cuối tuần và nghỉ lễ (đặc biệt Lễ 02/09).
    dt1 là ngày hạch toán Sao kê (Bank value date), dt2 là ngày hợp đồng CCTG.
    Quy chuẩn nghiệp vụ: Chấp nhận độ trễ thanh toán / chuyển tiền từ 0 đến +4 ngày (T+0..T+4),
    hoặc +5 ngày nếu trùng kỳ nghỉ Lễ Quốc Khánh 02/09 kết hợp cuối tuần.
    """
    if not dt1 or not dt2:
        return False, -999
    delta = (dt1 - dt2).days  # dt1 là ngày sao kê, dt2 là ngày giao dịch CCTG
    if 0 <= delta <= 4:
        return True, delta
    # Trường hợp nghỉ lễ 02/09 kết hợp cuối tuần (29/08 Thứ 7 -> 03/09 Thứ 5 = +5 ngày)
    if delta == 5 and dt2 == date(2026, 8, 29) and dt1 == date(2026, 9, 3):
        return True, delta
    return False, delta


def parse_date_str(val: str) -> str:
    """Chuẩn hóa chuỗi ngày tháng bất kỳ về định dạng hiển thị DD/MM/YYYY."""
    if not val or str(val).lower() in ("nan", "none", "null", ""):
        return ""
    val = str(val).strip()
    if len(val) == 8 and val.isdigit():
        return f"{val[6:8]}/{val[4:6]}/{val[0:4]}"
    for fmt in (
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
    ):
        try:
            dt = datetime.strptime(val, fmt)
            return dt.strftime("%d/%m/%Y")
        except Exception:
            pass
    return val


def clean_int_str(val: Any) -> str:
    """Làm sạch chuỗi số nguyên để không bị đuôi .0 (ví dụ '11199.0' -> '11199')."""
    if not val or str(val).lower() in ("nan", "none", "null", ""):
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def get_col(cols_list: List[str], candidates: List[str]) -> str:
    """Tìm tên cột đầu tiên khớp trong danh sách ứng viên (hỗ trợ đa schema)."""
    for c in candidates:
        if c in cols_list:
            return c
    return candidates[0]
