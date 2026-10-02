"""
Cấu Hình Hệ Thống & Đường Dẫn Dự Án ABBA CCTG
===========================================
Tập trung toàn bộ hằng số, regex, dung sai và đường dẫn thư mục chuẩn.
"""

import re
from pathlib import Path

# Thư mục gốc dự án
BASE_DIR = Path(__file__).resolve().parent.parent.parent

# Các phân vùng dữ liệu
DATA_DIR = BASE_DIR / "data"
INPUT_DIR = DATA_DIR / "input"
OUTPUT_DIR = DATA_DIR / "output"
ARCHIVE_DIR = DATA_DIR / "archive"
DOCS_DIR = BASE_DIR / "docs"

# Đảm bảo các thư mục tồn tại
for d in (INPUT_DIR, OUTPUT_DIR, ARCHIVE_DIR, DOCS_DIR):
    d.mkdir(parents=True, exist_ok=True)


def get_latest_data_file() -> Path:
    """Tự động tìm tệp dữ liệu CCTG mới nhất trong thư mục input."""
    # Ưu tiên các file csv/xlsx mới nhất theo tên
    candidates = list(INPUT_DIR.glob("Data_abba_*.csv")) + list(INPUT_DIR.glob("Data_abba_*.xlsx"))
    if candidates:
        # Sắp xếp theo tên để lấy ngày mới nhất (ví dụ 20261002 > 20260929)
        candidates.sort(key=lambda p: p.name, reverse=True)
        return candidates[0]
    
    # Fallback nếu tìm ở data/ cũ
    fallback = list(DATA_DIR.glob("Data_abba_*.csv")) + list(DATA_DIR.glob("Data_abba_*.xlsx"))
    if fallback:
        fallback.sort(key=lambda p: p.name, reverse=True)
        return fallback[0]
        
    return INPUT_DIR / "Data_abba_20261002.csv"


# Đường dẫn tệp mặc định
DEFAULT_DATA_PATH = get_latest_data_file()
DEFAULT_HD_PATH = INPUT_DIR / "BC Hoa don.xlsx"
DEFAULT_SK_PATH = INPUT_DIR / "saoke_TK43.xlsx"

# Đường dẫn tệp đầu ra chuẩn
OUTPUT_RECONCILE_PATH = OUTPUT_DIR / "Data_abba_Chuan_Va_DoiSoat.xlsx"
OUTPUT_CLIENT_EXCEL_PATH = OUTPUT_DIR / "Data_abba_20261002.xlsx"

# Tham số nghiệp vụ & Regex
AMOUNT_TOLERANCE = 1.0  # Dung sai tiền mặt ngân hàng (VND)
ROUNDING_TOLERANCE_BAN = 2000.0  # Dung sai làm tròn lãi bán lẻ từng sổ AZ (VND)
RE_FT = re.compile(r"(FT\d{12,16})", re.IGNORECASE)
RE_HD = re.compile(r"(CN[MB]-?[\w-]+)", re.IGNORECASE)
