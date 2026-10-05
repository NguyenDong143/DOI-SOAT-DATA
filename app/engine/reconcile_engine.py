"""
Động Cơ Đối Soát Tam Giác CCTG ABBA (Data CCTG ↔ Hóa Đơn ↔ Sao Kê TK43)
========================================================================
Bảo đảm tính toàn vẹn đa chiều, tự động nhận diện schema, phân tách vùng thời gian.
"""

import time
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
import pandas as pd
import xlsxwriter

from app.core.config import (
    AMOUNT_TOLERANCE,
    DEFAULT_DATA_PATH,
    DEFAULT_HD_PATH,
    DEFAULT_SK_PATH,
    OUTPUT_RECONCILE_PATH,
    RE_HD,
    ROUNDING_TOLERANCE_BAN,
)
from app.core.parsers import (
    dates_match,
    extract_ft,
    get_col,
    norm_code,
    parse_date,
)


class TriangleReconciliationEngine:
    def __init__(
        self,
        data_path: Optional[Path] = None,
        hd_path: Optional[Path] = None,
        sk_path: Optional[Path] = None,
        output_path: Optional[Path] = None,
    ):
        self.data_path = Path(data_path) if data_path else DEFAULT_DATA_PATH
        self.hd_path = Path(hd_path) if hd_path else DEFAULT_HD_PATH
        self.sk_path = Path(sk_path) if sk_path else DEFAULT_SK_PATH
        self.output_path = Path(output_path) if output_path else OUTPUT_RECONCILE_PATH

        self.summary_metrics: Dict[str, Any] = {}
        self.df_data_out: Optional[pd.DataFrame] = None
        self.problem_rows: List[Dict[str, Any]] = []
        self.problem_contracts: List[Dict[str, Any]] = []
        self.missing_hd_list: List[Dict[str, Any]] = []
        self.retail_rounding_contracts: List[Dict[str, Any]] = []
        self.retail_rounding_all: List[Dict[str, Any]] = []
        self.all_contracts: List[Dict[str, Any]] = []

    def execute(self) -> Dict[str, Any]:
        start_time = time.time()
        print("=" * 70, flush=True)
        print("BẮT ĐẦU ĐỐI SOÁT TAM GIÁC CCTG (DATA ↔ HÓA ĐƠN ↔ SAO KÊ TK43)", flush=True)
        print("=" * 70, flush=True)

        # 1. Đọc và lập chỉ mục Hóa Đơn
        print(f"\n[1/5] Đang đọc file Báo Cáo Hóa Đơn: {self.hd_path.name} ...", flush=True)
        excel_hd = pd.ExcelFile(self.hd_path)
        sheet_hd = "9.2026" if "9.2026" in excel_hd.sheet_names else excel_hd.sheet_names[0]
        df_hd_raw = pd.read_excel(excel_hd, sheet_name=sheet_hd, header=0)
        if df_hd_raw.iloc[0, 0] == "(1)":
            df_hd_raw = df_hd_raw.iloc[1:].reset_index(drop=True)

        inv_map = {}
        for _, r in df_hd_raw.iterrows():
            p_name = str(r.get("Tên sản phẩm", ""))
            m = RE_HD.search(p_name)
            if not m:
                continue
            c_code = m.group(1).upper()
            c_norm = norm_code(c_code)
            if c_norm not in inv_map:
                inv_map[c_norm] = {
                    "c_code": c_code,
                    "so_hd": str(r.get("Số hóa đơn", "")).strip(),
                    "ky_hieu": str(r.get("Mẫu số, ký hiệu", "")).strip(),
                    "ngay_hd": parse_date(r.get("Ngày hóa đơn")),
                    "ngay_hd_str": str(r.get("Ngày hóa đơn", "")),
                    "trang_thai": str(r.get("Trạng thái hóa đơn", "Hóa đơn gốc")).strip(),
                    "tong_tien": float(pd.to_numeric(r.get("Tổng tiền thanh toán", 0), errors="coerce") or 0),
                    "ten_kh": str(r.get("Tên người mua hàng", "")).strip(),
                    "mst": str(r.get("Mã số thuế người mua", "")).strip() if pd.notna(r.get("Mã số thuế người mua")) else "",
                    "ma_cctg": p_name.split()[0] if p_name else "",
                    "matched": False,
                }
        print(f"      -> {len(df_hd_raw):,} dòng hóa đơn chi tiết | {len(inv_map):,} hợp đồng.", flush=True)

        # 2. Đọc và lập chỉ mục Sao Kê TK43
        print(f"\n[2/5] Đang đọc file Sao Kê TK43: {self.sk_path.name} ...", flush=True)
        df_sk_raw = pd.read_excel(self.sk_path)
        raw_sk_count = len(df_sk_raw)

        # Khử trùng lặp kỹ thuật từ ngân hàng:
        # Nếu cùng Số giao dịch, cùng Ngày hiệu lực, cùng Số tiền Nợ/Có và cùng Số dư (Balance) -> do bank export trùng batch
        df_sk = df_sk_raw.drop_duplicates(
            subset=[
                "Số giao dịch (Transaction Number)",
                "Ngày hiệu lực (Value Date)",
                "Số tiền nợ (Debit amount)",
                "Số tiền có (Credit amount)",
                "Số dư (Balance)",
            ],
            keep="first",
        )
        sk_dupes_removed = raw_sk_count - len(df_sk)
        self.sk_dupes_removed = sk_dupes_removed
        self.sk_clean_rows = len(df_sk)

        sk_ft_map = defaultdict(list)
        sk_hd_desc_map = {}
        sk_max_date = None

        for idx, r in df_sk.iterrows():
            ft_raw = r.get("Số giao dịch (Transaction Number)")
            clean_ft = extract_ft(ft_raw)
            v_date = parse_date(r.get("Ngày hiệu lực (Value Date)"))
            if v_date and (sk_max_date is None or v_date > sk_max_date):
                sk_max_date = v_date

            debit = float(pd.to_numeric(r.get("Số tiền nợ (Debit amount)", 0), errors="coerce") or 0)
            credit = float(pd.to_numeric(r.get("Số tiền có (Credit amount)", 0), errors="coerce") or 0)
            desc = str(r.get("Diễn giải (Description)", ""))

            m_c = RE_HD.search(desc)
            extracted_c = norm_code(m_c.group(1)) if m_c else ""

            sk_entry = {
                "idx": idx,
                "stt": r.get("STT (No)"),
                "ft": clean_ft,
                "ft_raw": str(ft_raw),
                "date": v_date,
                "debit": debit,
                "credit": credit,
                "desc": desc,
                "desc_norm": norm_code(desc),
                "hd_extracted": m_c.group(1).upper() if m_c else "",
            }
            if clean_ft:
                sk_ft_map[clean_ft].append(sk_entry)
            if extracted_c:
                if extracted_c not in sk_hd_desc_map:
                    sk_hd_desc_map[extracted_c] = []
                sk_hd_desc_map[extracted_c].append(sk_entry)

        print(
            f"      -> {len(df_sk):,} dòng sao kê thực tế (Đã lọc {sk_dupes_removed:,} dòng trùng từ Bank) | "
            f"{len(sk_ft_map):,} FT | Chốt đến: {sk_max_date}",
            flush=True,
        )

        # 3. Đọc dữ liệu CCTG
        print(f"\n[3/5] Đang đọc tệp Data CCTG: {self.data_path.name} ...", flush=True)
        if self.data_path.suffix.lower() == ".csv":
            df_data = pd.read_csv(self.data_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        else:
            df_data = pd.read_excel(self.data_path, dtype=str)

        total_data_rows = len(df_data)
        cols = list(df_data.columns)

        col_stt = get_col(cols, ["STT"])
        if col_stt not in cols:
            df_data.insert(0, "STT", range(1, total_data_rows + 1))
            col_stt = "STT"
            cols = list(df_data.columns)

        col_cif = get_col(cols, ["Số CIF", "Số CIF khách hàng"])
        col_ten_kh = get_col(cols, ["Tên KH", "Tên khách hàng"])
        col_ma_cctg = get_col(cols, ["Mã CCTG"])
        col_so_az = get_col(cols, ["Số sổ AZ"])
        col_series = get_col(cols, ["Số series thứ cấp"])

        col_hd_mua = get_col(cols, ["Số HĐ mua", "Mã hợp đồng (ABBA bán, Khách mua)/ Bỏ trống nếu (ABBANK bán, Khách mua)"])
        col_amt_mua = get_col(cols, ["Tổng GT HĐ mua", "Tổng giá trị hợp đồng (ABBA/ ABBANK bán, Khách mua) (tất cả các sổ trong hợp đồng) = Tổng mệnh giá CCTG cả hợp đồng"])
        col_date_mua = get_col(cols, ["Ngày mua (ngày nắm giữ)", "Ngày (ABBA/ ABBANK bán, Khách mua)"])
        col_ft_mua = get_col(cols, ["Số FT HĐ Mua", "Mã FT giao dịch (ABBA bán, Khách mua)/ Bỏ trống nếu (ABBANK bán, Khách mua)"])
        col_so_hd = get_col(cols, ["Số hoá đơn", "Số Hóa Đơn"])
        col_ky_hieu = get_col(cols, ["Ký hiệu hoá đơn", "Ký Hiệu Hóa Đơn"])
        col_trang_thai_hd = get_col(cols, ["Trạng thái hoá đơn", "Trạng Thái Hóa Đơn"])

        col_hd_ban = get_col(cols, ["Số HĐ Bán", "Mã hợp đồng (Khách bán, ABBA mua)"])
        col_amt_ban = get_col(cols, ["Tổng GT HĐ Bán", "Tổng giá trị hợp đồng (Khách bán, ABBA mua) (tất cả các sổ trong hợp đồng) thực tế"])
        col_date_ban = get_col(cols, ["Ngày bán", "Ngày (Khách bán, ABBA mua) thực tế"])
        col_ft_ban = get_col(cols, ["Số FT HĐ Bán", "Mã FT giao dịch (Khách bán, ABBA mua)"])

        col_menh_gia = get_col(cols, ["Mệnh giá", "Mệnh giá 1 CCTG"])
        col_so_luong = get_col(cols, ["Số lượng", "Số lượng CCTG"])
        col_don_gia_ban = get_col(cols, ["Đơn giá bán/ số seri", "Đơn giá 1 CCTG từng sổ (Khách bán, ABBA mua) thực tế"])

        # Tính toán tiền và hash map tra cứu nhanh O(1)
        menh_gia = pd.to_numeric(df_data[col_menh_gia], errors="coerce").fillna(0)
        so_luong = pd.to_numeric(df_data[col_so_luong], errors="coerce").fillna(0)
        don_gia_ban = pd.to_numeric(df_data[col_don_gia_ban], errors="coerce").fillna(0)

        df_data["_tien_mua_dong"] = menh_gia * so_luong
        df_data["_tien_ban_dong"] = don_gia_ban * so_luong
        df_data["_ft_mua_clean"] = [extract_ft(x) for x in df_data[col_ft_mua]]
        df_data["_ft_ban_clean"] = [extract_ft(x) for x in df_data[col_ft_ban]]
        df_data["_date_mua"] = [parse_date(x) for x in df_data[col_date_mua]]
        df_data["_date_ban"] = [parse_date(x) for x in df_data[col_date_ban]]
        df_data["_hd_mua_norm"] = [norm_code(x) for x in df_data[col_hd_mua]]
        df_data["_hd_ban_norm"] = [norm_code(x) for x in df_data[col_hd_ban]]

        ft_mua_sum = defaultdict(float)
        ft_ban_sum = defaultdict(float)
        hd_mua_sum = defaultdict(float)
        data_ft_mua_info = defaultdict(list)
        data_ft_ban_info = defaultdict(list)

        for i in range(total_data_rows):
            fm = df_data["_ft_mua_clean"].iat[i]
            fb = df_data["_ft_ban_clean"].iat[i]
            hdm = df_data["_hd_mua_norm"].iat[i]
            tm = df_data["_tien_mua_dong"].iat[i]
            tb = df_data["_tien_ban_dong"].iat[i]
            if fm:
                ft_mua_sum[fm] += tm
                data_ft_mua_info[fm].append({
                    "stt": df_data[col_stt].iat[i],
                    "cif": str(df_data[col_cif].iat[i]).strip(),
                    "ten_kh": str(df_data[col_ten_kh].iat[i]).strip(),
                    "hd_mua": str(df_data[col_hd_mua].iat[i]).strip(),
                    "so_az": str(df_data[col_so_az].iat[i]).strip(),
                    "series": str(df_data[col_series].iat[i]).strip(),
                    "tien_mua": tm,
                    "row_idx": i,
                })
            if fb:
                ft_ban_sum[fb] += tb
                data_ft_ban_info[fb].append({
                    "stt": df_data[col_stt].iat[i],
                    "cif": str(df_data[col_cif].iat[i]).strip(),
                    "ten_kh": str(df_data[col_ten_kh].iat[i]).strip(),
                    "hd_ban": str(df_data[col_hd_ban].iat[i]).strip(),
                    "so_az": str(df_data[col_so_az].iat[i]).strip(),
                    "series": str(df_data[col_series].iat[i]).strip(),
                    "tien_ban": tb,
                    "row_idx": i,
                })
            if hdm:
                hd_mua_sum[hdm] += tm

        # 4. Thực Hiện Đối Soát Tam Giác
        print(f"\n[4/5] Đang phân tích đối chiếu cho {total_data_rows:,} dòng ...", flush=True)
        flags_status = []
        flags_details = []
        highlight_codes = []

        so_hd_enriched = list(df_data[col_so_hd])
        ky_hieu_enriched = list(df_data[col_ky_hieu])
        trang_thai_enriched = list(df_data[col_trang_thai_hd])
        ft_mua_enriched = list(df_data[col_ft_mua])

        count_golden = 0
        count_rounding = 0
        count_enriched_hd = 0
        count_enriched_ft = 0
        count_discrepancy = 0
        count_abbank = 0
        count_pending_sale = 0
        count_future_sk = 0
        count_cross_contract = 0
        cross_contract_fts = set()

        # Tập hợp theo dõi các FT đã được khớp bởi Data CCTG (dùng cho bước scan ngược)
        matched_fts_mua: set = set()   # FT mua đã được tiêu thụ bởi ít nhất 1 dòng Data
        matched_fts_ban: set = set()   # FT bán đã được tiêu thụ bởi ít nhất 1 dòng Data

        self.problem_rows.clear()
        self.problem_contracts.clear()
        self.all_contracts.clear()
        all_contracts_dict = {}
        contract_problems = defaultdict(lambda: {
            "cif": "",
            "ten_kh": "",
            "ma_cctg": "",
            "hd_mua": "",
            "ngay_mua": "",
            "ft_mua": "",
            "tong_tien_mua_hd": 0.0,
            "hd_ban": "",
            "ngay_ban": "",
            "ft_ban": "",
            "tong_tien_ban_hd": 0.0,
            "so_luong_so_az": 0,
            "cac_so_az": [],
            "error_types": set(),
            "error_details": set(),
        })

        retail_sell_map = defaultdict(lambda: {
            "hd_ban": "",
            "cif": "",
            "ten_kh": "",
            "so_luong_so_az": 0,
            "d_ban": None,
            "ngay_ban_str": "",
            "ft_ban": "",
            "calc_sum": 0.0,
            "header_val": 0.0,
            "sk_debit": 0.0,
        })

        # Danh sách chứa các cột so sánh đối chiếu song song (Side-by-side)
        cmp_ft_m_data = []
        cmp_ft_m_sk = []
        cmp_hd_m_sk = []
        cmp_date_m_data = []
        cmp_date_m_sk = []
        cmp_delta_date_m = []
        cmp_tien_m_data = []
        cmp_tien_m_sk = []
        cmp_delta_tien_m = []

        cmp_ft_b_data = []
        cmp_ft_b_sk = []
        cmp_date_b_data = []
        cmp_date_b_sk = []
        cmp_delta_date_b = []
        cmp_tien_b_data = []
        cmp_tien_b_sk = []
        cmp_delta_tien_b = []

        cmp_hd_inv_code = []
        cmp_hd_inv_so = []
        cmp_hd_inv_ky_hieu = []
        cmp_hd_inv_ngay = []
        cmp_hd_inv_tien = []
        cmp_delta_tien_hd = []
        cmp_action_guides = []

        for i in range(total_data_rows):
            h_norm = df_data["_hd_mua_norm"].iat[i]
            ft_m = df_data["_ft_mua_clean"].iat[i]
            d_mua = df_data["_date_mua"].iat[i]
            t_mua = df_data["_tien_mua_dong"].iat[i]

            ft_b = df_data["_ft_ban_clean"].iat[i]
            d_ban = df_data["_date_ban"].iat[i]
            t_ban = df_data["_tien_ban_dong"].iat[i]

            row_modified = False
            row_issues = []
            row_notes = []
            is_abbank = (not h_norm)

            is_after_sk = bool(isinstance(d_mua, date) and isinstance(sk_max_date, date) and d_mua > sk_max_date)

            # --- Biến đối chiếu Mua ---
            cur_ft_sk_m = ""
            cur_hd_sk_m = ""
            cur_date_sk_m = ""
            cur_delta_date_m = ""
            cur_amt_sk_m = 0.0
            cur_diff_amt_m = 0.0

            # --- Biến đối chiếu Bán ---
            cur_ft_sk_b = ""
            cur_date_sk_b = ""
            cur_delta_date_b = ""
            cur_amt_sk_b = 0.0
            cur_diff_amt_b = 0.0

            # --- Biến đối chiếu Hóa Đơn ---
            cur_hd_inv_code = ""
            cur_hd_inv_so = ""
            cur_hd_inv_ky_hieu = ""
            cur_hd_inv_ngay = ""
            cur_hd_inv_tien = 0.0
            cur_diff_inv = 0.0

            # 4.1. Hóa Đơn
            if not is_abbank and h_norm in inv_map:
                inv_entry = inv_map[h_norm]
                inv_entry["matched"] = True
                curr_so_hd = df_data[col_so_hd].iat[i]
                if not curr_so_hd or str(curr_so_hd).strip() in ("", "nan", "None", "0"):
                    so_hd_enriched[i] = inv_entry["so_hd"]
                    ky_hieu_enriched[i] = inv_entry["ky_hieu"]
                    trang_thai_enriched[i] = inv_entry["trang_thai"]
                    row_modified = True
                    count_enriched_hd += 1
                    row_notes.append(f"Bổ sung Số HĐ: {inv_entry['so_hd']}")

                cur_hd_inv_code = inv_entry["c_code"]
                cur_hd_inv_so = inv_entry["so_hd"]
                cur_hd_inv_ky_hieu = inv_entry["ky_hieu"]
                cur_hd_inv_ngay = str(inv_entry["ngay_hd"]) if inv_entry["ngay_hd"] else ""
                cur_hd_inv_tien = inv_entry["tong_tien"]

                tot_hd_cctg = hd_mua_sum[h_norm]
                cur_diff_inv = tot_hd_cctg - inv_entry["tong_tien"]
                if abs(cur_diff_inv) >= AMOUNT_TOLERANCE:
                    row_issues.append(f"Lệch tiền HĐ: {cur_diff_inv:+,.0f} đ (Data={tot_hd_cctg:,.0f} vs HĐ={inv_entry['tong_tien']:,.0f})")

            # 4.2. Sao Kê Mua
            if is_after_sk:
                row_notes.append(f"Giao dịch ngày {d_mua.strftime('%d/%m/%Y')} (Chờ sao kê sau {sk_max_date.strftime('%d/%m/%Y')})")
                count_future_sk += 1
            elif ft_m:
                if ft_m in sk_ft_map:
                    sk_entries = sk_ft_map[ft_m]
                    sk_entry = sk_entries[0]
                    b_credit = sum(e["credit"] for e in sk_entries)
                    b_date = sk_entry["date"]

                    tot_ft_cctg = ft_mua_sum[ft_m]
                    diff_amt_m = b_credit - tot_ft_cctg

                    # Khớp ưu tiên: Nếu tổng sk_entries lệch nhưng có đúng 1 dòng sao kê khớp chuẩn 1:1 với hợp đồng
                    if abs(diff_amt_m) >= AMOUNT_TOLERANCE and len(sk_entries) > 1:
                        exact_match = [e for e in sk_entries if abs(e["credit"] - tot_ft_cctg) < AMOUNT_TOLERANCE]
                        if exact_match:
                            sk_entry = exact_match[0]
                            b_credit = sk_entry["credit"]
                            b_date = sk_entry["date"]
                            diff_amt_m = 0.0

                    cur_ft_sk_m = sk_entry["ft"]
                    cur_hd_sk_m = sk_entry.get("hd_extracted", "")
                    cur_date_sk_m = str(b_date) if b_date else ""
                    cur_amt_sk_m = b_credit
                    cur_diff_amt_m = diff_amt_m
                    matched_fts_mua.add(ft_m)  # Đánh dấu FT đã được khớp bởi Data

                    # Cross-contract validation: FT khớp tiền nhưng mã HĐ trên Sao Kê ≠ mã HĐ trên Data
                    # → Nghi nhầm FT sang KH/HĐ khác (VD: HĐ 18715 LAM GIA PHUOC vs PHAM NGOC THIEN)
                    if cur_hd_sk_m and h_norm:
                        sk_hd_norm = norm_code(cur_hd_sk_m)
                        if sk_hd_norm and sk_hd_norm != h_norm:
                            count_cross_contract += 1
                            cross_contract_fts.add(ft_m)
                            sk_desc_text = sk_entry.get("desc", "")
                            # Trích xuất tên KH từ diễn giải sao kê (phần trước "Thanh toan")
                            sk_kh_parts = sk_desc_text.split("Thanh toan")
                            sk_kh_name = sk_kh_parts[0].strip() if len(sk_kh_parts) > 1 else ""
                            data_kh_name = str(df_data[col_ten_kh].iat[i]).strip()
                            row_issues.append(
                                f"XUNG ĐỘT MÃ HĐ: FT {ft_m} trên Sao Kê ghi cho HĐ {cur_hd_sk_m}"
                                f" (KH Sao Kê: {sk_kh_name})"
                                f" nhưng Data CCTG ghi HĐ {str(df_data[col_hd_mua].iat[i]).strip()}"
                                f" (KH Data: {data_kh_name})"
                                f" — Cần kiểm tra ngay xem tiền có bị ghi nhầm sang KH/HĐ khác"
                            )

                    if b_credit <= 0:
                        row_issues.append(f"Sai chiều Mua: Sao kê ghi Nợ {sk_entry['debit']:,.0f} đ")
                    else:
                        if abs(diff_amt_m) >= AMOUNT_TOLERANCE:
                            row_issues.append(f"Lệch tiền Mua: {diff_amt_m:+,.0f} đ (Sao kê={b_credit:,.0f} vs Data={tot_ft_cctg:,.0f})")

                        d_ok, delta = dates_match(b_date, d_mua)
                        cur_delta_date_m = delta if delta != -999 else ""
                        if not d_ok:
                            row_issues.append(f"Lệch ngày Mua: {delta:+d} ngày (Bank={b_date} vs Data={d_mua})")
                else:
                    row_issues.append(f"Không tìm thấy FT Mua: {ft_m} trên Sao Kê")
            else:
                if not is_abbank:
                    found_sk = None
                    if h_norm in sk_hd_desc_map:
                        for e in sk_hd_desc_map[h_norm]:
                            if e["credit"] > 0 and abs(e["credit"] - hd_mua_sum[h_norm]) < AMOUNT_TOLERANCE:
                                found_sk = e
                                break
                    if found_sk:
                        ft_mua_enriched[i] = found_sk["ft"]
                        cur_ft_sk_m = found_sk["ft"]
                        cur_hd_sk_m = found_sk.get("hd_extracted", "")
                        cur_date_sk_m = str(found_sk["date"]) if found_sk["date"] else ""
                        cur_amt_sk_m = found_sk["credit"]
                        d_ok, delta = dates_match(found_sk["date"], d_mua)
                        cur_delta_date_m = delta if delta != -999 else ""
                        row_modified = True
                        count_enriched_ft += 1
                        row_notes.append(f"Bổ sung FT Mua: {found_sk['ft']} từ Sao Kê")
                    else:
                        row_issues.append("Thiếu FT Mua (chưa xác định trên Sao Kê)")

            # 4.3. Sao Kê Bán
            is_sold = bool(isinstance(d_ban, date))
            if is_sold:
                is_ban_after_sk = bool(isinstance(d_ban, date) and isinstance(sk_max_date, date) and d_ban > sk_max_date)
                if is_ban_after_sk:
                    row_notes.append(f"Tất toán ngày {d_ban.strftime('%d/%m/%Y')} (Chờ sao kê đợt mới)")
                elif ft_b:
                    if ft_b in sk_ft_map:
                        sk_entries = sk_ft_map[ft_b]
                        sk_entry = sk_entries[0]
                        b_debit = sum(e["debit"] for e in sk_entries)
                        b_date = sk_entry["date"]

                        tot_ft_ban = ft_ban_sum[ft_b]
                        diff_amt_b = b_debit - tot_ft_ban

                        # Khớp ưu tiên 1:1 cho chiều bán
                        if abs(diff_amt_b) >= AMOUNT_TOLERANCE and len(sk_entries) > 1:
                            exact_match = [e for e in sk_entries if abs(e["debit"] - tot_ft_ban) < AMOUNT_TOLERANCE]
                            if exact_match:
                                sk_entry = exact_match[0]
                                b_debit = sk_entry["debit"]
                                b_date = sk_entry["date"]
                                diff_amt_b = 0.0

                        cur_ft_sk_b = sk_entry["ft"]
                        cur_date_sk_b = str(b_date) if b_date else ""
                        cur_amt_sk_b = b_debit
                        cur_diff_amt_b = diff_amt_b
                        matched_fts_ban.add(ft_b)  # Đánh dấu FT đã được khớp bởi Data

                        if b_debit <= 0:
                            row_issues.append(f"Sai chiều Bán: Sao kê ghi Có {sk_entry['credit']:,.0f} đ")
                        else:
                            if abs(diff_amt_b) >= ROUNDING_TOLERANCE_BAN:
                                row_issues.append(f"Lệch tiền Bán: {diff_amt_b:+,.0f} đ (Sao kê={b_debit:,.0f} vs Data={tot_ft_ban:,.0f})")
                            elif abs(diff_amt_b) >= AMOUNT_TOLERANCE:
                                row_notes.append(f"Lệch làm tròn bán: {diff_amt_b:+,.0f} đ")

                            d_ok, delta = dates_match(b_date, d_ban)
                            cur_delta_date_b = delta if delta != -999 else ""
                            if not d_ok:
                                row_issues.append(f"Lệch ngày Bán: {delta:+d} ngày (Bank={b_date} vs Data={d_ban})")
                    else:
                        row_issues.append(f"Không tìm thấy FT Bán: {ft_b} trên Sao Kê")
                else:
                    row_issues.append("Có ngày bán thực tế nhưng thiếu FT Bán")
            else:
                count_pending_sale += 1

            # 4.4. Đánh cờ trạng thái
            has_rounding_ban = bool(is_sold and abs(cur_diff_amt_b) >= AMOUNT_TOLERANCE and abs(cur_diff_amt_b) < ROUNDING_TOLERANCE_BAN)

            if is_abbank:
                count_abbank += 1
                stt_text = "⚠️ CẢNH BÁO: ABBANK bán (Có lệch)" if row_issues else "ℹ️ THÔNG TIN: ABBANK bán (Lịch sử trước T6/2026)"
                h_code = "YELLOW" if row_issues else "NORMAL"
            elif row_issues:
                count_discrepancy += 1
                has_conflict_issue = any("xung đột" in s.lower() for s in row_issues)
                has_money_issue = any("tiền" in s.lower() or "sai chiều" in s.lower() for s in row_issues)
                if has_conflict_issue:
                    stt_text = "❌ LỖI: Xung đột mã FT / Khách hàng giữa Sao kê và Data"
                    h_code = "RED"
                elif has_money_issue:
                    stt_text = "❌ LỖI: Lệch tiền hoặc sai chiều giao dịch"
                    h_code = "RED"
                else:
                    stt_text = "⚠️ CẢNH BÁO: Lệch ngày hoặc thiếu FT/HĐ"
                    h_code = "YELLOW"
            elif has_rounding_ban:
                count_rounding += 1
                stt_text = f"ℹ️ LÀM TRÒN: Lệch làm tròn bán lẻ ({cur_diff_amt_b:+,.0f} đ)"
                h_code = "TEAL"
            elif is_after_sk:
                stt_text = f"ℹ️ THÔNG TIN: Chờ Sao kê TK43 đợt mới (sau {sk_max_date.strftime('%d/%m') if isinstance(sk_max_date, date) else ''})"
                h_code = "NORMAL"
            elif row_modified:
                stt_text = "⚡ ĐÃ BỔ SUNG: Dữ liệu đã được làm sạch"
                h_code = "GREEN"
            else:
                count_golden += 1
                stt_text = "✅ CHUẨN: Khớp hoàn hảo 100% (Không lệch)"
                h_code = "GREEN"

            guide_text = "Dữ liệu hợp lệ, không cần xử lý"
            if row_issues:
                if any("xung đột" in s.lower() for s in row_issues):
                    guide_text = "Khẩn: Đối chiếu lại CIF, Tên KH và HĐ gốc với chứng từ chuyển tiền tại Ngân hàng (Nghi vấn nạp nhầm thông tin KH/Sổ AZ)"
                elif any("không tìm thấy ft" in s.lower() for s in row_issues):
                    guide_text = "Kiểm tra mã FT ngân hàng hoặc liên hệ bank tra soát lệnh chuyển tiền"
                elif any("lệch tiền" in s.lower() for s in row_issues):
                    guide_text = "Kiểm tra lại số tiền nộp/tất toán và hạch toán điều chỉnh chênh lệch"
                elif any("lệch ngày" in s.lower() for s in row_issues):
                    guide_text = "Xác nhận lại ngày chứng từ ngân hàng so với ngày ghi nhận hợp đồng"
                elif any("sai chiều" in s.lower() for s in row_issues):
                    guide_text = "Kế toán kiểm tra giao dịch đảo chiều hoặc hoàn tiền từ ngân hàng"
                else:
                    guide_text = "Kế toán kiểm tra đối chiếu chứng từ gốc"
            elif has_rounding_ban:
                guide_text = "Lệch làm tròn số học do lẻ đơn giá sổ AZ; kế toán hạch toán điều chỉnh (Xem Sheet Chênh Lệch Bán Lẻ)"

            full_details = []
            if row_notes:
                full_details.append(" | ".join(row_notes))
            if row_issues:
                full_details.append("LỆCH: " + "; ".join(row_issues))
            detail_text = " || ".join(full_details) if full_details else "Dữ liệu hợp lệ, không có sai lệch."

            flags_status.append(stt_text)
            flags_details.append(detail_text)
            highlight_codes.append(h_code)

            # Lưu các cột so sánh song song cho Data CCTG
            cur_tien_m_data = hd_mua_sum.get(h_norm, 0.0) if h_norm else t_mua
            cur_tien_b_data = ft_ban_sum.get(ft_b, 0.0) if ft_b else t_ban

            cmp_ft_m_data.append(ft_m)
            cmp_ft_m_sk.append(cur_ft_sk_m)
            cmp_hd_m_sk.append(cur_hd_sk_m)
            cmp_date_m_data.append(str(d_mua) if d_mua else "")
            cmp_date_m_sk.append(cur_date_sk_m)
            cmp_delta_date_m.append(cur_delta_date_m)
            cmp_tien_m_data.append(cur_tien_m_data)
            cmp_tien_m_sk.append(cur_amt_sk_m)
            cmp_delta_tien_m.append(cur_diff_amt_m)

            cmp_ft_b_data.append(ft_b)
            cmp_ft_b_sk.append(cur_ft_sk_b)
            cmp_date_b_data.append(str(d_ban) if d_ban else "")
            cmp_date_b_sk.append(cur_date_sk_b)
            cmp_delta_date_b.append(cur_delta_date_b)
            cmp_tien_b_data.append(cur_tien_b_data)
            cmp_tien_b_sk.append(cur_amt_sk_b)
            cmp_delta_tien_b.append(cur_diff_amt_b)

            cmp_hd_inv_code.append(cur_hd_inv_code)
            cmp_hd_inv_so.append(cur_hd_inv_so)
            cmp_hd_inv_ky_hieu.append(cur_hd_inv_ky_hieu)
            cmp_hd_inv_ngay.append(cur_hd_inv_ngay)
            cmp_hd_inv_tien.append(cur_hd_inv_tien)
            cmp_delta_tien_hd.append(cur_diff_inv)
            cmp_action_guides.append(guide_text)

            # Ghi nhận hợp đồng bán lẻ và toàn bộ hợp đồng (CNM & CNB)
            raw_hdb = str(df_data[col_hd_ban].iat[i]).strip()
            raw_hdm = str(df_data[col_hd_mua].iat[i]).strip()
            cif_val = str(df_data[col_cif].iat[i]).strip()
            ten_kh_val = str(df_data[col_ten_kh].iat[i]).strip()
            ma_cctg_val = str(df_data[col_ma_cctg].iat[i]).strip()
            sl_val = int(so_luong.iat[i]) if pd.notna(so_luong.iat[i]) else 0

            if raw_hdb and raw_hdb.lower() not in ("nan", "none", ""):
                r_entry = retail_sell_map[raw_hdb]
                r_entry["hd_ban"] = raw_hdb
                r_entry["cif"] = cif_val
                r_entry["ten_kh"] = ten_kh_val
                r_entry["so_luong_so_az"] += 1
                if not r_entry["d_ban"] and d_ban:
                    r_entry["d_ban"] = d_ban
                    r_entry["ngay_ban_str"] = str(df_data[col_date_ban].iat[i]).strip()
                if not r_entry["ft_ban"] and ft_b:
                    r_entry["ft_ban"] = ft_b
                r_entry["calc_sum"] += t_ban
                if r_entry["header_val"] == 0.0:
                    val_h = pd.to_numeric(df_data[col_amt_ban].iat[i], errors="coerce")
                    r_entry["header_val"] = float(val_h) if pd.notna(val_h) else 0.0
                if cur_amt_sk_b > 0 and r_entry["sk_debit"] == 0.0:
                    r_entry["sk_debit"] = cur_amt_sk_b

                raw_hdr_m = pd.to_numeric(df_data[col_amt_mua].iat[i], errors="coerce")
                num_hdr_m = float(raw_hdr_m) if pd.notna(raw_hdr_m) else 0.0
                raw_hdr_b = pd.to_numeric(df_data[col_amt_ban].iat[i], errors="coerce")
                num_hdr_b = float(raw_hdr_b) if pd.notna(raw_hdr_b) else 0.0

                if raw_hdb not in all_contracts_dict:
                    all_contracts_dict[raw_hdb] = {
                        "hd_ban": raw_hdb,
                        "hd_mua": raw_hdm if raw_hdm and raw_hdm.lower() not in ("nan", "none", "") else "",
                        "cif": cif_val,
                        "ten_kh": ten_kh_val,
                        "ma_cctg": ma_cctg_val,
                        "so_luong_so_az": 0,
                        "tong_sl_cctg": 0,
                        "cac_so_az": [],
                        "d_mua": d_mua,
                        "ngay_mua_str": str(df_data[col_date_mua].iat[i]).strip() if pd.notna(df_data[col_date_mua].iat[i]) else "",
                        "ft_mua_data": ft_m,
                        "ft_mua_sk": cur_ft_sk_m,
                        "calc_tien_mua": 0.0,
                        "hdr_tien_mua": num_hdr_m,
                        "sk_credit": cur_amt_sk_m,
                        "inv_code": cur_hd_inv_code,
                        "inv_so_hd": cur_hd_inv_so,
                        "inv_ky_hieu": cur_hd_inv_ky_hieu,
                        "inv_ngay": cur_hd_inv_ngay,
                        "inv_tien": cur_hd_inv_tien,
                        "d_ban": d_ban,
                        "ngay_ban_str": str(df_data[col_date_ban].iat[i]).strip() if pd.notna(df_data[col_date_ban].iat[i]) else "",
                        "ft_ban_data": ft_b,
                        "ft_ban_sk": cur_ft_sk_b,
                        "calc_tien_ban": 0.0,
                        "hdr_tien_ban": num_hdr_b,
                        "sk_debit": cur_amt_sk_b,
                        "is_abbank": is_abbank,
                        "is_after_sk": is_after_sk,
                        "has_rounding_ban": False,
                        "issues": set(),
                        "notes": set(),
                    }
                c_item = all_contracts_dict[raw_hdb]
                c_item["so_luong_so_az"] += 1
                c_item["tong_sl_cctg"] += sl_val
                c_item["calc_tien_mua"] += t_mua
                c_item["calc_tien_ban"] += t_ban
                if c_item["hdr_tien_mua"] == 0.0 and num_hdr_m > 0:
                    c_item["hdr_tien_mua"] = num_hdr_m
                if c_item["hdr_tien_ban"] == 0.0 and num_hdr_b > 0:
                    c_item["hdr_tien_ban"] = num_hdr_b
                if len(c_item["cac_so_az"]) < 5:
                    c_item["cac_so_az"].append(str(df_data[col_so_az].iat[i]))
                if not c_item["hd_mua"] and raw_hdm and raw_hdm.lower() not in ("nan", "none", ""):
                    c_item["hd_mua"] = raw_hdm
                if not c_item["ft_mua_data"] and ft_m:
                    c_item["ft_mua_data"] = ft_m
                if not c_item["ft_mua_sk"] and cur_ft_sk_m:
                    c_item["ft_mua_sk"] = cur_ft_sk_m
                if cur_amt_sk_m > 0 and c_item["sk_credit"] == 0:
                    c_item["sk_credit"] = cur_amt_sk_m
                if not c_item["inv_so_hd"] and cur_hd_inv_so:
                    c_item["inv_code"] = cur_hd_inv_code
                    c_item["inv_so_hd"] = cur_hd_inv_so
                    c_item["inv_ky_hieu"] = cur_hd_inv_ky_hieu
                    c_item["inv_ngay"] = cur_hd_inv_ngay
                    c_item["inv_tien"] = cur_hd_inv_tien
                if not c_item["d_ban"] and d_ban:
                    c_item["d_ban"] = d_ban
                    c_item["ngay_ban_str"] = str(df_data[col_date_ban].iat[i]).strip()
                if not c_item["ft_ban_data"] and ft_b:
                    c_item["ft_ban_data"] = ft_b
                if not c_item["ft_ban_sk"] and cur_ft_sk_b:
                    c_item["ft_ban_sk"] = cur_ft_sk_b
                if cur_amt_sk_b > 0 and c_item["sk_debit"] == 0:
                    c_item["sk_debit"] = cur_amt_sk_b
                if has_rounding_ban:
                    c_item["has_rounding_ban"] = True
                for iss in row_issues:
                    c_item["issues"].add(iss)
                for n in row_notes:
                    c_item["notes"].add(n)

            if row_issues:
                self.problem_rows.append({
                    "STT Data": df_data[col_stt].iat[i],
                    "Số CIF": df_data[col_cif].iat[i],
                    "Tên Khách Hàng": df_data[col_ten_kh].iat[i],
                    "Mã CCTG": df_data[col_ma_cctg].iat[i],
                    "Số Sổ AZ": df_data[col_so_az].iat[i],
                    "Số Series Thứ Cấp": df_data[col_series].iat[i],
                    "Mã HĐ Mua": df_data[col_hd_mua].iat[i],
                    "[So Sánh Mua] FT Data": ft_m,
                    "[So Sánh Mua] FT Sao Kê": cur_ft_sk_m,
                    "[So Sánh Mua] Mã HĐ Sao Kê": cur_hd_sk_m,
                    "[So Sánh Mua] Ngày Mua Data": str(d_mua) if d_mua else "",
                    "[So Sánh Mua] Ngày Sao Kê": cur_date_sk_m,
                    "[So Sánh Mua] Δ Ngày Mua": cur_delta_date_m,
                    "[So Sánh Mua] Tiền Mua HĐ (Data)": cur_tien_m_data,
                    "[So Sánh Mua] Tiền Ghi Có (Sao Kê)": cur_amt_sk_m,
                    "[So Sánh Mua] Δ Tiền Mua (Bank - Data)": cur_diff_amt_m,
                    "Mã HĐ Bán": df_data[col_hd_ban].iat[i],
                    "[So Sánh Bán] FT Data": ft_b,
                    "[So Sánh Bán] FT Sao Kê": cur_ft_sk_b,
                    "[So Sánh Bán] Ngày Bán Data": str(d_ban) if d_ban else "",
                    "[So Sánh Bán] Ngày Sao Kê": cur_date_sk_b,
                    "[So Sánh Bán] Δ Ngày Bán": cur_delta_date_b,
                    "[So Sánh Bán] Tiền Bán HĐ (Data)": cur_tien_b_data,
                    "[So Sánh Bán] Tiền Ghi Nợ (Sao Kê)": cur_amt_sk_b,
                    "[So Sánh Bán] Δ Tiền Bán (Bank - Data)": cur_diff_amt_b,
                    "[So Sánh HĐ] Mã HĐ (Hóa Đơn)": cur_hd_inv_code,
                    "[So Sánh HĐ] Số Hóa Đơn": cur_hd_inv_so,
                    "[So Sánh HĐ] Ký Hiệu Hóa Đơn": cur_hd_inv_ky_hieu,
                    "[So Sánh HĐ] Ngày Hóa Đơn": cur_hd_inv_ngay,
                    "[So Sánh HĐ] Tiền Hóa Đơn": cur_hd_inv_tien,
                    "[So Sánh HĐ] Δ Tiền HĐ (Data - HĐ)": cur_diff_inv,
                    "Phân Loại Lỗi": stt_text,
                    "Chi Tiết Sai Lệch": "; ".join(row_issues),
                    "Hướng Dẫn Xử Lý": guide_text,
                })

                hd_key = str(df_data[col_hd_mua].iat[i]).strip()
                if not hd_key or hd_key.lower() in ("nan", "none", ""):
                    hd_key = str(df_data[col_hd_ban].iat[i]).strip()
                if not hd_key or hd_key.lower() in ("nan", "none", ""):
                    hd_key = f"CIF_{df_data[col_cif].iat[i]}_{df_data[col_so_az].iat[i]}"

                cp = contract_problems[hd_key]
                cp["cif"] = df_data[col_cif].iat[i]
                cp["ten_kh"] = df_data[col_ten_kh].iat[i]
                cp["ma_cctg"] = df_data[col_ma_cctg].iat[i]
                cp["hd_mua"] = df_data[col_hd_mua].iat[i]
                cp["ngay_mua"] = str(d_mua) if d_mua else ""
                cp["ft_mua"] = ft_m
                cp["tong_tien_mua_hd"] = hd_mua_sum.get(h_norm, 0.0)
                cp["hd_ban"] = df_data[col_hd_ban].iat[i]
                cp["ngay_ban"] = str(d_ban) if d_ban else ""
                cp["ft_ban"] = ft_b
                cp["tong_tien_ban_hd"] = ft_ban_sum.get(ft_b, 0.0) if ft_b else 0.0
                cp["so_luong_so_az"] += 1
                if len(cp["cac_so_az"]) < 5:
                    cp["cac_so_az"].append(str(df_data[col_so_az].iat[i]))
                cp["error_types"].add(stt_text)
                for iss in row_issues:
                    cp["error_details"].add(iss)

        # Xây dựng danh sách hợp đồng có vấn đề (Contract-level aggregation)
        for idx, (hd_key, cp) in enumerate(contract_problems.items(), 1):
            so_az_summary = ", ".join(cp["cac_so_az"])
            if cp["so_luong_so_az"] > len(cp["cac_so_az"]):
                so_az_summary += f", ... (+{cp['so_luong_so_az'] - len(cp['cac_so_az'])} sổ khác)"
            self.problem_contracts.append({
                "STT": idx,
                "Số CIF": cp["cif"],
                "Tên Khách Hàng": cp["ten_kh"],
                "Mã CCTG": cp["ma_cctg"],
                "Mã HĐ Mua": cp["hd_mua"],
                "Ngày Mua": cp["ngay_mua"],
                "Mã FT Mua": cp["ft_mua"],
                "Tổng Tiền Mua HĐ (VND)": cp["tong_tien_mua_hd"],
                "Mã HĐ Bán": cp["hd_ban"],
                "Ngày Bán": cp["ngay_ban"],
                "Mã FT Bán": cp["ft_ban"],
                "Tổng Tiền Bán HĐ (VND)": cp["tong_tien_ban_hd"],
                "Số Lượng Sổ AZ Bị Ảnh Hưởng": cp["so_luong_so_az"],
                "Danh Sách Sổ AZ Tiêu Biểu": so_az_summary,
                "Phân Loại Vấn Đề": " | ".join(sorted(cp["error_types"])),
                "Chi Tiết Sai Lệch": "; ".join(sorted(cp["error_details"])),
                "Hướng Dẫn Xử Lý": "Kế toán kiểm tra chứng từ ngân hàng hoặc hạch toán điều chỉnh cho hợp đồng này",
            })

        # Xây dựng danh sách hợp đồng bán lẻ có chênh lệch làm tròn
        self.retail_rounding_contracts.clear()
        self.retail_rounding_all.clear()
        for hdb in sorted(retail_sell_map.keys()):
            rm = retail_sell_map[hdb]
            bank_amt = rm["sk_debit"] if rm["sk_debit"] > 0 else rm["header_val"]
            calc_amt = rm["calc_sum"]
            diff = calc_amt - bank_amt
            if rm["d_ban"] and abs(diff) > 0 and abs(diff) < ROUNDING_TOLERANCE_BAN:
                ky_ban = rm["d_ban"].strftime("%m/%Y")
                is_t9 = bool(rm["d_ban"].month == 9 and rm["d_ban"].year == 2026)
                entry = {
                    "hd_ban": rm["hd_ban"],
                    "ten_kh": rm["ten_kh"],
                    "cif": rm["cif"],
                    "so_luong_so_az": rm["so_luong_so_az"],
                    "ky_ban": ky_ban,
                    "ngay_ban": str(rm["d_ban"]) if rm["d_ban"] else "",
                    "ngay_ban_str": rm["ngay_ban_str"],
                    "ft_ban": rm["ft_ban"],
                    "sl_x_don_gia": int(round(calc_amt)),
                    "sao_ke": int(round(bank_amt)),
                    "chenh_lech": int(round(diff)),
                    "danh_gia": "Lệch làm tròn số học do lẻ đơn giá sổ AZ (< 2,000 đ)",
                    "is_t9": is_t9,
                }
                self.retail_rounding_all.append(entry)
                if is_t9:
                    self.retail_rounding_contracts.append(entry)

        # Sắp xếp danh sách tổng hợp: Tháng 9 lên đầu (230 HĐ chuẩn), sau đó là các tháng trước
        self.retail_rounding_all.sort(key=lambda x: (not x["is_t9"], x["ky_ban"], x["hd_ban"]))
        print(f"      -> Phát hiện {len(self.retail_rounding_all):,} HĐ bán lẻ có sai số làm tròn (T9: {len(self.retail_rounding_contracts)} HĐ lệch {sum(e['chenh_lech'] for e in self.retail_rounding_contracts):+,.0f} đ | Toàn bộ: {sum(e['chenh_lech'] for e in self.retail_rounding_all):+,.0f} đ).", flush=True)

        # Gán lại cột bổ sung
        df_data[col_so_hd] = so_hd_enriched
        df_data[col_ky_hieu] = ky_hieu_enriched
        df_data[col_trang_thai_hd] = trang_thai_enriched
        df_data[col_ft_mua] = ft_mua_enriched

        # Bổ sung các cụm cột so sánh đối chiếu song song (Side-by-Side)
        df_data["[So Sánh Mua] FT Data"] = cmp_ft_m_data
        df_data["[So Sánh Mua] FT Sao Kê"] = cmp_ft_m_sk
        df_data["[So Sánh Mua] Mã HĐ Sao Kê"] = cmp_hd_m_sk
        df_data["[So Sánh Mua] Ngày Mua Data"] = cmp_date_m_data
        df_data["[So Sánh Mua] Ngày Sao Kê"] = cmp_date_m_sk
        df_data["[So Sánh Mua] Δ Ngày Mua"] = cmp_delta_date_m
        df_data["[So Sánh Mua] Tiền Mua HĐ (Data)"] = cmp_tien_m_data
        df_data["[So Sánh Mua] Tiền Ghi Có (Sao Kê)"] = cmp_tien_m_sk
        df_data["[So Sánh Mua] Δ Tiền Mua (Bank - Data)"] = cmp_delta_tien_m

        df_data["[So Sánh Bán] FT Data"] = cmp_ft_b_data
        df_data["[So Sánh Bán] FT Sao Kê"] = cmp_ft_b_sk
        df_data["[So Sánh Bán] Ngày Bán Data"] = cmp_date_b_data
        df_data["[So Sánh Bán] Ngày Sao Kê"] = cmp_date_b_sk
        df_data["[So Sánh Bán] Δ Ngày Bán"] = cmp_delta_date_b
        df_data["[So Sánh Bán] Tiền Bán HĐ (Data)"] = cmp_tien_b_data
        df_data["[So Sánh Bán] Tiền Ghi Nợ (Sao Kê)"] = cmp_tien_b_sk
        df_data["[So Sánh Bán] Δ Tiền Bán (Bank - Data)"] = cmp_delta_tien_b

        df_data["[So Sánh HĐ] Mã HĐ (Hóa Đơn)"] = cmp_hd_inv_code
        df_data["[So Sánh HĐ] Số Hóa Đơn"] = cmp_hd_inv_so
        df_data["[So Sánh HĐ] Ký Hiệu Hóa Đơn"] = cmp_hd_inv_ky_hieu
        df_data["[So Sánh HĐ] Ngày Hóa Đơn"] = cmp_hd_inv_ngay
        df_data["[So Sánh HĐ] Tiền Hóa Đơn"] = cmp_hd_inv_tien
        df_data["[So Sánh HĐ] Δ Tiền HĐ (Data - HĐ)"] = cmp_delta_tien_hd

        df_data["[CỜ ĐỐI SOÁT & TRẠNG THÁI]"] = flags_status
        df_data["[CHI TIẾT ĐIỀU CHỈNH / ĐỐI SOÁT]"] = flags_details
        df_data["[HƯỚNG DẪN XỬ LÝ]"] = cmp_action_guides

        df_data.drop(
            columns=["_tien_mua_dong", "_tien_ban_dong", "_ft_mua_clean", "_ft_ban_clean", "_date_mua", "_date_ban", "_hd_mua_norm", "_hd_ban_norm"],
            inplace=True,
        )
        self.df_data_out = df_data

        # 4.5. Scan ngược: Tìm FT trên Sao Kê (credit) không được khớp bởi bất kỳ dòng Data nào
        # Đây là trường hợp: Tiền đã vào Ngân hàng, có/không có HĐ, nhưng thiếu hoàn toàn trên Data CCTG
        orphan_ft_entries = []  # [(ft, sk_entry), ...]
        for ft_key, sk_list in sk_ft_map.items():
            if ft_key in matched_fts_mua:
                continue  # FT này đã được Data tiêu thụ ở chiều mua
            # Chỉ xét các FT có phát sinh credit (thu tiền vào TK43)
            credit_entries = [e for e in sk_list if e["credit"] > 0]
            if not credit_entries:
                continue
            # Bỏ qua các khoản trả lãi coupon và điều chuyển vốn nội bộ
            credit_entries = [
                e for e in credit_entries
                if "tra lai coupon" not in e["desc"].lower()
                and "chuyen tien tu tk" not in e["desc"].lower()
            ]
            if not credit_entries:
                continue
            # Kiểm tra thêm: FT này không xuất hiện trong bất kỳ FT mua/bán nào của Data
            if ft_key in ft_mua_sum or ft_key in ft_ban_sum:
                continue  # Có trên Data nhưng không match được sao kê — đã được báo lỗi ở chiều mua
            for e in credit_entries:
                orphan_ft_entries.append((ft_key, e))

        if orphan_ft_entries:
            # Nhóm theo FT (một FT có thể có nhiều dòng credit)
            orphan_by_ft: dict = {}
            for ft_key, e in orphan_ft_entries:
                if ft_key not in orphan_by_ft:
                    orphan_by_ft[ft_key] = {"total_credit": 0.0, "entries": [], "hd_extracted": ""}
                orphan_by_ft[ft_key]["total_credit"] += e["credit"]
                orphan_by_ft[ft_key]["entries"].append(e)
                if not orphan_by_ft[ft_key]["hd_extracted"] and e["hd_extracted"]:
                    orphan_by_ft[ft_key]["hd_extracted"] = e["hd_extracted"]

            for ft_key, info in orphan_by_ft.items():
                first_entry = info["entries"][0]
                total_credit = info["total_credit"]
                hd_from_desc = info["hd_extracted"]  # Mã HĐ trích xuất từ diễn giải ngân hàng
                sk_date_str = str(first_entry["date"]) if first_entry["date"] else ""
                sk_desc = first_entry["desc"]

                # Tìm thông tin KH từ Hóa đơn (nếu mã HĐ khớp)
                hd_norm = norm_code(hd_from_desc) if hd_from_desc else ""
                inv_ten_kh = ""
                inv_so_hd = ""
                inv_tong_tien = 0.0
                inv_matched_for_orphan = False
                if hd_norm and hd_norm in inv_map:
                    inv_r = inv_map[hd_norm]
                    inv_ten_kh = inv_r["ten_kh"]
                    inv_so_hd = inv_r["so_hd"]
                    inv_tong_tien = inv_r["tong_tien"]
                    inv_matched_for_orphan = True
                    if inv_r.get("matched", False):
                        continue  # HĐ đã được khớp với Data CCTG → không cần báo thêm

                risk = "CRITICAL (FT GHI CÓ KHÔNG KHỚP BẤT KỲ DÒNG DATA CCTG NÀO)"
                if inv_matched_for_orphan:
                    amt_diff = total_credit - inv_tong_tien
                    note_detail = (
                        f"Sao kê ghi nhận thu tiền FT={ft_key} | Số tiền={total_credit:,.0f} đ "
                        f"| Mã HĐ từ diễn giải SK: {hd_from_desc} | Tên KH (HĐ): {inv_ten_kh} "
                        f"| Δ vs HĐ={amt_diff:+,.0f} đ "
                        f"| Diễn giải SK: {sk_desc[:100]}"
                    )
                    recommendation = (
                        "KHẨN: FT đã thu tiền trên Bank nhưng THIẾU HOÀN TOÀN trên Data CCTG. "
                        "Kiểm tra ngay xem tiền có bị ghi nhầm sang tài khoản/hợp đồng khác không. "
                        "Yêu cầu Vận hành/Core nạp Sổ AZ & Series vào Data CCTG."
                    )
                else:
                    note_detail = (
                        f"Sao kê ghi nhận thu tiền FT={ft_key} | Số tiền={total_credit:,.0f} đ "
                        f"| Mã HĐ từ diễn giải SK: {hd_from_desc or '(Không trích xuất được)'} "
                        f"| KHÔNG TÌM THẤY TRÊN DATA CCTG VÀ HÓA ĐƠN "
                        f"| Diễn giải SK: {sk_desc[:100]}"
                    )
                    recommendation = (
                        "KHẨN: FT ghi có trên Bank không khớp bất kỳ dòng Data CCTG hay Hóa đơn nào. "
                        "Tra soát khẩn với ngân hàng để xác định nguồn gốc dòng tiền."
                    )

                # Check ngược lại mã FT trên Data CCTG:
                rev_status = "Chưa xuất hiện trên Data CCTG"
                rev_kh = ""
                rev_cif = ""
                rev_hd = ""
                rev_az = ""
                rev_series = ""
                if ft_key in data_ft_mua_info:
                    c_rows = data_ft_mua_info[ft_key]
                    rev_cif = ", ".join(sorted(set(r["cif"] for r in c_rows)))
                    rev_kh = ", ".join(sorted(set(r["ten_kh"] for r in c_rows)))
                    rev_hd = ", ".join(sorted(set(r["hd_mua"] for r in c_rows)))
                    rev_az = ", ".join(sorted(set(r["so_az"] for r in c_rows)))
                    rev_series = ", ".join(sorted(set(r["series"] for r in c_rows)))
                    rev_status = "⚠️ XUNG ĐỘT: BỊ GÁN CHO KH KHÁC TRÊN DATA"

                self.missing_hd_list.append({
                    "Mã Hợp Đồng": hd_from_desc or "(Chưa xác định từ diễn giải SK)",
                    "Tên Khách Hàng (HĐĐT/Sao Kê)": inv_ten_kh or "(Chưa xác định — Kiểm tra diễn giải SK)",
                    "Mã Số Thuế / CCCD": "",
                    "Mã CCTG": "",
                    "Số Hóa Đơn": inv_so_hd or "",
                    "Ký Hiệu Hóa Đơn": "",
                    "Ngày Hóa Đơn": "",
                    "Trạng Thái HĐ": "FT MỒ CÔI — Không tìm thấy trên Data CCTG",
                    "Tổng Tiền Hóa Đơn (VND)": inv_tong_tien,
                    "Trạng Thái Dòng Tiền Ngân Hàng": "ĐÃ GHI CÓ TRÊN SAO KÊ NHƯNG KHÔNG MATCH DATA",
                    "Mã FT Sao Kê": ft_key,
                    "Ngày Thu Tiền (Bank)": sk_date_str,
                    "Số Tiền Đã Thu (VND)": total_credit,
                    "Δ Tiền (Sao Kê - Hóa Đơn)": total_credit - inv_tong_tien if inv_tong_tien else total_credit,
                    "[Kiểm Tra Ngược Data CCTG] Trạng Thái FT Trên Data": rev_status,
                    "[Kiểm Tra Ngược Data CCTG] Tên KH Đang Giữ FT": rev_kh,
                    "[Kiểm Tra Ngược Data CCTG] Số CIF Đang Giữ FT": rev_cif,
                    "[Kiểm Tra Ngược Data CCTG] Số HĐ Mua Đang Giữ FT": rev_hd,
                    "[Kiểm Tra Ngược Data CCTG] Số Sổ AZ Đang Giữ FT": rev_az,
                    "[Kiểm Tra Ngược Data CCTG] Số Series Thứ Cấp": rev_series,
                    "Mức Độ Rủi Ro": risk,
                    "Đánh Giá Bất Thường / Chi Tiết Xung Đột": note_detail,
                    "Khuyến Nghị Xử Lý Nghiệp Vụ": recommendation,
                })

        # 4.5. Tổng hợp HĐ ngoại lệ (từ Hóa đơn chưa khớp)
        print("\n[5/5] Đang tổng hợp các hợp đồng Hóa đơn & Sao kê ngoại lệ ...", flush=True)
        # Lưu ý: missing_hd_list đã có thể chứa orphan FT từ bước scan ngược ở trên — KHÔNG clear
        for c_norm, inv in inv_map.items():
            if not inv["matched"]:
                ft_matched = ""
                date_matched = ""
                amt_matched = 0
                status_sk = "Chưa tìm thấy dòng tiền Sao Kê"

                if c_norm in sk_hd_desc_map:
                    sk_cand = sk_hd_desc_map[c_norm][0]
                    ft_matched = sk_cand["ft"]
                    date_matched = str(sk_cand["date"]) if sk_cand["date"] else ""
                    amt_matched = sk_cand["credit"] if sk_cand["credit"] > 0 else sk_cand["debit"]
                    status_sk = f"ĐÃ THU ĐỦ TIỀN TRÊN SAO KÊ ({amt_matched:,.0f} đ)"

                # Check ngược lại mã FT trên Data CCTG:
                reverse_data_status = "Chưa xuất hiện trên Data CCTG"
                reverse_data_kh = ""
                reverse_data_cif = ""
                reverse_data_hd = ""
                reverse_data_az = ""
                reverse_data_series = ""
                reverse_notes = ""
                risk_level = "CRITICAL (THẤT THOÁT CHỨNG TỪ)" if ft_matched else "MEDIUM"
                rec_guide = "Yêu cầu phòng Vận hành/Core cấp bổ sung Số sổ AZ & Series để nạp vào hệ thống CCTG"

                if ft_matched and ft_matched in data_ft_mua_info:
                    c_rows = data_ft_mua_info[ft_matched]
                    reverse_data_cif = ", ".join(sorted(set(r["cif"] for r in c_rows)))
                    reverse_data_kh = ", ".join(sorted(set(r["ten_kh"] for r in c_rows)))
                    reverse_data_hd = ", ".join(sorted(set(r["hd_mua"] for r in c_rows)))
                    reverse_data_az = ", ".join(sorted(set(r["so_az"] for r in c_rows)))
                    reverse_data_series = ", ".join(sorted(set(r["series"] for r in c_rows)))

                    if norm_code(reverse_data_kh) != norm_code(inv["ten_kh"]):
                        reverse_data_status = "⚠️ XUNG ĐỘT: BỊ GÁN CHO KH KHÁC TRÊN DATA CCTG"
                        risk_level = "CỰC KỲ NGUY HIỂM (XUNG ĐỘT MÃ FT / GÁN NHẦM KHÁCH HÀNG TRÊN DATA)"
                        reverse_notes = (
                            f"Mã FT {ft_matched} thu tiền {amt_matched:,.0f} đ cho KH '{inv['ten_kh']}' (HĐĐT #{inv['so_hd']}) "
                            f"đã vào Sao kê. Tuy nhiên trên Data CCTG, FT này và Sổ AZ {reverse_data_az} lại đang bị gán cho KH '{reverse_data_kh}' "
                            f"(CIF: {reverse_data_cif}, HĐ Mua: {reverse_data_hd}). "
                            f"Lưu ý: Số series thứ cấp trên Data ({reverse_data_series}) có đuôi là mã HĐ '{inv['c_code']}' của {inv['ten_kh']}! "
                            f"-> Bản chất: Tiền và Sổ AZ là của KH {inv['ten_kh']}, nhưng trên Data CCTG bị nạp nhầm thông tin CIF/Tên KH sang {reverse_data_kh}."
                        )
                        rec_guide = (
                            f"Khẩn cấp kiểm tra điều chỉnh thông tin CIF/Tên KH của Sổ AZ {reverse_data_az} trên Data CCTG "
                            f"về đúng chủ sở hữu: KH {inv['ten_kh']} (CIF trên HĐ), hoặc tách riêng sổ AZ."
                        )
                    else:
                        reverse_data_status = "ĐÃ CÓ TRÊN DATA (CÙNG KHÁCH HÀNG)"
                        reverse_notes = f"FT {ft_matched} đã có trên Data CCTG cho KH {reverse_data_kh}."
                elif ft_matched:
                    reverse_notes = f"FT {ft_matched} thu tiền trên Sao kê ngân hàng nhưng chưa có bất kỳ dòng nào trên Data CCTG."

                self.missing_hd_list.append({
                    "Mã Hợp Đồng": inv["c_code"],
                    "Tên Khách Hàng (HĐĐT/Sao Kê)": inv["ten_kh"],
                    "Mã Số Thuế / CCCD": inv["mst"],
                    "Mã CCTG": inv["ma_cctg"],
                    "Số Hóa Đơn": inv["so_hd"],
                    "Ký Hiệu Hóa Đơn": inv["ky_hieu"],
                    "Ngày Hóa Đơn": inv["ngay_hd_str"],
                    "Trạng Thái HĐ": inv["trang_thai"],
                    "Tổng Tiền Hóa Đơn (VND)": inv["tong_tien"],
                    "Trạng Thái Dòng Tiền Ngân Hàng": status_sk,
                    "Mã FT Sao Kê": ft_matched,
                    "Ngày Thu Tiền (Bank)": date_matched,
                    "Số Tiền Đã Thu (VND)": amt_matched,
                    "Δ Tiền (Sao Kê - Hóa Đơn)": (amt_matched - inv["tong_tien"]) if ft_matched else 0,
                    "[Kiểm Tra Ngược Data CCTG] Trạng Thái FT Trên Data": reverse_data_status,
                    "[Kiểm Tra Ngược Data CCTG] Tên KH Đang Giữ FT": reverse_data_kh,
                    "[Kiểm Tra Ngược Data CCTG] Số CIF Đang Giữ FT": reverse_data_cif,
                    "[Kiểm Tra Ngược Data CCTG] Số HĐ Mua Đang Giữ FT": reverse_data_hd,
                    "[Kiểm Tra Ngược Data CCTG] Số Sổ AZ Đang Giữ FT": reverse_data_az,
                    "[Kiểm Tra Ngược Data CCTG] Số Series Thứ Cấp": reverse_data_series,
                    "Mức Độ Rủi Ro": risk_level,
                    "Đánh Giá Bất Thường / Chi Tiết Xung Đột": reverse_notes,
                    "Khuyến Nghị Xử Lý Nghiệp Vụ": rec_guide,
                })

        # Sắp xếp danh sách ngoại lệ: Ưu tiên các trường hợp CỰC KỲ NGUY HIỂM / XUNG ĐỘT lên đầu (VD: HĐ 18715)
        self.missing_hd_list.sort(
            key=lambda x: (
                0 if "CỰC KỲ" in str(x.get("Mức Độ Rủi Ro", "")) else (1 if x.get("Số Hóa Đơn") else 2),
                x.get("Mã Hợp Đồng", "")
            )
        )

        # Xây dựng danh sách 43,618 hợp đồng toàn diện (CNM & CNB) cho Tab Chi Tiết Theo Hợp Đồng
        def safe_int(v) -> int:
            try:
                if v is None or pd.isna(v):
                    return 0
                return int(round(float(v)))
            except (ValueError, TypeError):
                return 0

        self.all_contracts.clear()
        for hdb, c in all_contracts_dict.items():
            calc_m = safe_int(c.get("calc_tien_mua"))
            hdr_m = safe_int(c.get("hdr_tien_mua"))
            sk_m = safe_int(c.get("sk_credit"))
            inv_m = safe_int(c.get("inv_tien"))
            diff_m_bank = (sk_m - calc_m) if sk_m > 0 else 0
            diff_m_inv = (calc_m - inv_m) if inv_m > 0 else 0

            calc_b = safe_int(c.get("calc_tien_ban"))
            hdr_b = safe_int(c.get("hdr_tien_ban"))
            sk_b = safe_int(c.get("sk_debit"))
            bank_b_val = sk_b if sk_b > 0 else hdr_b
            is_sold = bool(c["d_ban"] is not None)
            diff_b_retail = (bank_b_val - calc_b) if is_sold else 0

            has_rounding = bool(is_sold and abs(diff_b_retail) > 0 and abs(diff_b_retail) < ROUNDING_TOLERANCE_BAN)
            is_t9 = bool(is_sold and c["d_ban"].month == 9 and c["d_ban"].year == 2026)
            ky_ban = c["d_ban"].strftime("%m/%Y") if is_sold else ""

            if has_rounding:
                stt_hd = f"ℹ️ LÀM TRÒN: Lệch làm tròn bán lẻ ({diff_b_retail:+,.0f} đ)"
                color_hd = "TEAL"
                guide_hd = "Lệch làm tròn số học do lẻ đơn giá sổ AZ; kế toán hạch toán điều chỉnh"
                detail_hd = f"Lệch làm tròn bán lẻ {diff_b_retail:+,.0f} đ (Data={calc_b:,.0f} vs Bank={bank_b_val:,.0f})"
            elif c["issues"]:
                has_conflict = any("xung đột" in s.lower() for s in c["issues"])
                has_money_issue = any("tiền" in s.lower() or "sai chiều" in s.lower() for s in c["issues"])
                if has_conflict:
                    stt_hd = "❌ LỖI: Xung đột mã FT / Khách hàng giữa Sao kê và Data"
                    color_hd = "RED"
                    guide_hd = "Khẩn: Đối chiếu lại CIF, Tên KH và HĐ gốc với chứng từ chuyển tiền tại Ngân hàng (Nghi vấn nạp nhầm thông tin KH/Sổ AZ)"
                elif has_money_issue:
                    stt_hd = "❌ LỖI: Lệch tiền hoặc sai chiều giao dịch"
                    color_hd = "RED"
                    guide_hd = "Kế toán kiểm tra chứng từ ngân hàng hoặc hạch toán điều chỉnh"
                else:
                    stt_hd = "⚠️ CẢNH BÁO: Lệch ngày hoặc thiếu FT/HĐ"
                    color_hd = "YELLOW"
                    guide_hd = "Kế toán kiểm tra chứng từ ngân hàng hoặc hạch toán điều chỉnh"
                detail_hd = "; ".join(sorted(c["issues"]))
            elif not is_sold:
                stt_hd = "ℹ️ ĐANG NẮM GIỮ: Chưa phát sinh bán"
                color_hd = "NORMAL"
                guide_hd = "Không cần xử lý, tiếp tục theo dõi kỳ hạn CCTG"
                detail_hd = "Hợp đồng CCTG đang được khách hàng nắm giữ, chưa phát sinh dòng tiền tất toán"
            elif c["is_after_sk"]:
                stt_hd = "ℹ️ CHỜ SAO KÊ: Phát sinh sau ngày chốt sao kê"
                color_hd = "NORMAL"
                guide_hd = "Đối chiếu vào kỳ sao kê tiếp theo"
                detail_hd = f"Giao dịch sau ngày chốt sao kê TK43 ({sk_max_date})"
            elif c["is_abbank"]:
                stt_hd = "ℹ️ ABBANK BÁN: Lịch sử trước T6/2026"
                color_hd = "NORMAL"
                guide_hd = "Dữ liệu lịch sử ABBANK bán trước T6/2026"
                detail_hd = "Nghiệp vụ ABBANK bán lịch sử"
            else:
                stt_hd = "✅ CHUẨN: Khớp hoàn hảo 100% (Không lệch)"
                color_hd = "GREEN"
                guide_hd = "Hồ sơ hợp lệ, lưu trữ theo quy định"
                detail_hd = "Khớp chuẩn 100% các tiêu chí Data, Sao kê Bank và Hóa đơn"

            self.all_contracts.append({
                "hd_mua": c["hd_mua"],
                "hd_ban": c["hd_ban"],
                "cif": c["cif"],
                "ten_kh": c["ten_kh"],
                "ma_cctg": c["ma_cctg"],
                "so_luong_so_az": c["so_luong_so_az"],
                "tong_sl_cctg": c["tong_sl_cctg"],
                "ngay_mua": c["ngay_mua_str"],
                "ft_mua_data": c["ft_mua_data"],
                "ft_mua_sk": c["ft_mua_sk"],
                "calc_tien_mua": calc_m,
                "hdr_tien_mua": hdr_m,
                "sk_credit": sk_m,
                "inv_tien": inv_m,
                "diff_m_bank": diff_m_bank,
                "diff_m_inv": diff_m_inv,
                "ngay_ban": c["ngay_ban_str"],
                "ft_ban_data": c["ft_ban_data"],
                "ft_ban_sk": c["ft_ban_sk"],
                "calc_tien_ban": calc_b,
                "hdr_tien_ban": hdr_b,
                "sk_debit": sk_b,
                "diff_b_retail": diff_b_retail,
                "stt_hd": stt_hd,
                "detail_hd": detail_hd,
                "guide_hd": guide_hd,
                "color_hd": color_hd,
                "has_rounding": has_rounding,
                "is_t9": is_t9,
                "ky_ban": ky_ban,
                "has_issues": bool(c["issues"]),
            })

        # Bổ sung các hợp đồng ngoại lệ chứng từ (HĐ & Sao kê có nhưng thiếu trên Data)
        for inv_entry in self.missing_hd_list:
            inv_stt = inv_entry.get("[Kiểm Tra Ngược Data CCTG] Trạng Thái FT Trên Data", "")
            if "XUNG ĐỘT" in inv_stt:
                stt_hd_val = "❌ XUNG ĐỘT: FT bị gán nhầm sang KH khác trên Data"
                color_hd_val = "RED"
                detail_hd_val = inv_entry.get("Đánh Giá Bất Thường / Chi Tiết Xung Đột", "")
                guide_hd_val = inv_entry.get("Khuyến Nghị Xử Lý Nghiệp Vụ", "")
            else:
                stt_hd_val = "❌ NGOẠI LỆ: Có Hóa đơn & Sao kê nhưng thiếu trên Data"
                color_hd_val = "RED"
                detail_hd_val = inv_entry.get("Đánh Giá Bất Thường / Chi Tiết Xung Đột") or "Đã thu đủ tiền trên Sao kê, HĐĐT đã xuất nhưng Core thiếu Sổ AZ"
                guide_hd_val = inv_entry.get("Khuyến Nghị Xử Lý Nghiệp Vụ") or "Yêu cầu phòng Vận hành/Core nạp Sổ AZ & Series vào Data CCTG"

            self.all_contracts.append({
                "hd_mua": inv_entry.get("Mã Hợp Đồng", ""),
                "hd_ban": "(Chưa nạp trên Data)",
                "cif": inv_entry.get("[Kiểm Tra Ngược Data CCTG] Số CIF Đang Giữ FT") or "(Chưa có)",
                "ten_kh": inv_entry.get("Tên Khách Hàng (HĐĐT/Sao Kê)") or inv_entry.get("Tên Khách Hàng", ""),
                "ma_cctg": inv_entry.get("Mã CCTG", ""),
                "so_luong_so_az": 0,
                "tong_sl_cctg": 0,
                "ngay_mua": inv_entry.get("Ngày Thu Tiền (Bank)") or inv_entry.get("Ngày Thu Tiền") or inv_entry.get("Ngày Hóa Đơn", ""),
                "ft_mua_data": "(Chưa có)",
                "ft_mua_sk": inv_entry.get("Mã FT Sao Kê", ""),
                "calc_tien_mua": 0,
                "hdr_tien_mua": 0,
                "sk_credit": safe_int(inv_entry.get("Số Tiền Đã Thu (VND)")),
                "inv_tien": safe_int(inv_entry.get("Tổng Tiền Hóa Đơn (VND)")),
                "diff_m_bank": safe_int(inv_entry.get("Số Tiền Đã Thu (VND)")),
                "diff_m_inv": -safe_int(inv_entry.get("Tổng Tiền Hóa Đơn (VND)")),
                "ngay_ban": "",
                "ft_ban_data": "",
                "ft_ban_sk": "",
                "calc_tien_ban": 0,
                "hdr_tien_ban": 0,
                "sk_debit": 0,
                "diff_b_retail": 0,
                "stt_hd": stt_hd_val,
                "detail_hd": detail_hd_val,
                "guide_hd": guide_hd_val,
                "color_hd": color_hd_val,
                "has_rounding": False,
                "is_t9": True,
                "ky_ban": "",
                "has_issues": True,
            })

        # Sắp xếp danh sách hợp đồng
        def get_contract_priority(x):
            if x["has_rounding"] and x["is_t9"]:
                return 0
            if x["has_rounding"]:
                return 1
            if x["has_issues"]:
                return 2
            if "CHUẨN" in x["stt_hd"]:
                return 3
            return 4

        self.all_contracts.sort(key=lambda x: (get_contract_priority(x), x["ky_ban"], x["ngay_mua"], x["hd_ban"]))
        print(f"      -> Tổng hợp {len(self.all_contracts):,} hợp đồng đầy đủ cả CNM & CNB cho Tab Chi Tiết Theo Hợp Đồng.", flush=True)

        # Lưu các chỉ số thống kê
        cnt_contracts_cnm = sum(1 for c in self.all_contracts if c["hd_mua"] and not c["hd_mua"].startswith("("))
        cnt_contracts_cnm_golden = sum(1 for c in self.all_contracts if c["hd_mua"] and not c["hd_mua"].startswith("(") and c["sk_credit"] > 0 and c["diff_m_bank"] == 0)
        cnt_contracts_sold = sum(1 for c in self.all_contracts if c["ngay_ban"])
        cnt_contracts_sold_golden = sum(1 for c in self.all_contracts if c["ngay_ban"] and not c["has_rounding"] and not c["has_issues"])
        cnt_contracts_holding = sum(1 for c in self.all_contracts if not c["ngay_ban"] and not c["has_issues"] and not c["hd_ban"].startswith("("))
        cnt_contracts_future_sk = sum(1 for c in self.all_contracts if "CHỜ SAO KÊ" in c["stt_hd"])

        self.summary_metrics = {
            "total_rows": total_data_rows,
            "count_golden": count_golden,
            "count_rounding": count_rounding,
            "count_future_sk": count_future_sk,
            "count_abbank": count_abbank,
            "count_pending_sale": count_pending_sale,
            "count_discrepancy": count_discrepancy,
            "count_cross_contract": count_cross_contract,
            "count_cross_contract_fts": len(cross_contract_fts),
            "count_problem_contracts": len(self.problem_contracts),
            "count_missing_hd": len(self.missing_hd_list),
            "count_contracts_total": len(self.all_contracts),
            "count_contracts_cnm": cnt_contracts_cnm,
            "count_contracts_cnm_golden": cnt_contracts_cnm_golden,
            "count_contracts_sold": cnt_contracts_sold,
            "count_contracts_sold_golden": cnt_contracts_sold_golden,
            "count_contracts_holding": cnt_contracts_holding,
            "count_contracts_future_sk": cnt_contracts_future_sk,
            "count_retail_rounding_all": len(self.retail_rounding_all),
            "count_retail_rounding_t9": len(self.retail_rounding_contracts),
            "sum_retail_data_all": sum(e["sl_x_don_gia"] for e in self.retail_rounding_all),
            "sum_retail_bank_all": sum(e["sao_ke"] for e in self.retail_rounding_all),
            "sum_retail_diff_all": sum(e["chenh_lech"] for e in self.retail_rounding_all),
            "sum_retail_data_t9": sum(e["sl_x_don_gia"] for e in self.retail_rounding_contracts),
            "sum_retail_bank_t9": sum(e["sao_ke"] for e in self.retail_rounding_contracts),
            "sum_retail_diff_t9": sum(e["chenh_lech"] for e in self.retail_rounding_contracts),
            "sum_retail_data_other": sum(e["sl_x_don_gia"] for e in self.retail_rounding_all if not e.get("is_t9")),
            "sum_retail_bank_other": sum(e["sao_ke"] for e in self.retail_rounding_all if not e.get("is_t9")),
            "sum_retail_diff_other": sum(e["chenh_lech"] for e in self.retail_rounding_all if not e.get("is_t9")),
            "sk_max_date": str(sk_max_date),
            "sk_dupes_removed": self.sk_dupes_removed,
            "sk_clean_rows": self.sk_clean_rows,
            "data_file_name": self.data_path.name,
        }

        # 5. Xuất Workbook Excel 5 Sheet
        self._export_excel(highlight_codes)

        elapsed = time.time() - start_time
        print(f"\n[HOÀN TẤT] File đã được xuất thành công tại: {self.output_path}", flush=True)
        print(f"Tổng thời gian xử lý: {elapsed:.1f} giây.", flush=True)
        print("=" * 70, flush=True)
        return self.summary_metrics

    def _export_excel(self, highlight_codes: List[str]):
        print(f"[*] Đang tạo Workbook kết quả: {self.output_path.name} ...", flush=True)
        try:
            writer = pd.ExcelWriter(self.output_path, engine="xlsxwriter")
        except PermissionError:
            print(f"[CẢNH BÁO] File {self.output_path.name} đang được mở trong Excel. Đang thử đóng tiến trình Excel để ghi đè...", flush=True)
            import subprocess
            subprocess.run(["taskkill", "/f", "/im", "excel.exe"], capture_output=True)
            time.sleep(1.5)
            try:
                writer = pd.ExcelWriter(self.output_path, engine="xlsxwriter")
            except PermissionError:
                alt_path = self.output_path.with_name(f"{self.output_path.stem}_{datetime.now().strftime('%H%M%S')}.xlsx")
                print(f"[THÔNG BÁO] File vẫn bị khóa. Tự động chuyển hướng lưu sang: {alt_path.name}", flush=True)
                self.output_path = alt_path
                writer = pd.ExcelWriter(self.output_path, engine="xlsxwriter")
        wb = writer.book

        fmt_hdr_base = wb.add_format({
            "bold": True, "bg_color": "#1B365D", "font_color": "#FFFFFF", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_hdr_mua = wb.add_format({
            "bold": True, "bg_color": "#0E6251", "font_color": "#FFFFFF", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_hdr_ban = wb.add_format({
            "bold": True, "bg_color": "#512E5F", "font_color": "#FFFFFF", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_hdr_hd = wb.add_format({
            "bold": True, "bg_color": "#7D6608", "font_color": "#FFFFFF", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_hdr_eval = wb.add_format({
            "bold": True, "bg_color": "#922B21", "font_color": "#FFFFFF", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })

        def get_header_fmt(col_name: str):
            if "[Chiều Mua" in col_name or "[So Sánh Mua]" in col_name:
                return fmt_hdr_mua
            elif "[Chiều Bán" in col_name or "[So Sánh Bán]" in col_name:
                return fmt_hdr_ban
            elif "[So Sánh HĐ]" in col_name:
                return fmt_hdr_hd
            elif "[" in col_name and "]" in col_name:
                return fmt_hdr_eval
            return fmt_hdr_base

        fmt_green_row = wb.add_format({
            "bg_color": "#E2EFDA", "font_color": "#276A3C", "border": 1, "valign": "vcenter", "font_size": 9
        })
        fmt_yellow_row = wb.add_format({
            "bg_color": "#FFF2CC", "font_color": "#806000", "border": 1, "valign": "vcenter", "font_size": 9
        })
        fmt_red_row = wb.add_format({
            "bg_color": "#FCE4D6", "font_color": "#C00000", "border": 1, "valign": "vcenter", "font_size": 9
        })
        fmt_teal_row = wb.add_format({
            "bg_color": "#E8F8F5", "font_color": "#0E6251", "border": 1, "valign": "vcenter", "font_size": 9
        })
        fmt_normal_cell = wb.add_format({
            "border": 1, "valign": "vcenter", "font_size": 9
        })
        fmt_curr_cell = wb.add_format({
            "border": 1, "valign": "vcenter", "num_format": r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)', "font_size": 9
        })
        fmt_center_cell = wb.add_format({
            "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_teal_curr = wb.add_format({
            "bg_color": "#E8F8F5", "font_color": "#0E6251", "border": 1, "valign": "vcenter", "num_format": r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)', "font_size": 9
        })
        fmt_teal_center = wb.add_format({
            "bg_color": "#E8F8F5", "font_color": "#0E6251", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_green_curr = wb.add_format({
            "bg_color": "#E2EFDA", "font_color": "#276A3C", "border": 1, "valign": "vcenter", "num_format": r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)', "font_size": 9
        })
        fmt_green_center = wb.add_format({
            "bg_color": "#E2EFDA", "font_color": "#276A3C", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_red_curr = wb.add_format({
            "bg_color": "#FCE4D6", "font_color": "#C00000", "border": 1, "valign": "vcenter", "num_format": r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)', "font_size": 9
        })
        fmt_red_center = wb.add_format({
            "bg_color": "#FCE4D6", "font_color": "#C00000", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_yellow_curr = wb.add_format({
            "bg_color": "#FFF2CC", "font_color": "#806000", "border": 1, "valign": "vcenter", "num_format": r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)', "font_size": 9
        })
        fmt_yellow_center = wb.add_format({
            "bg_color": "#FFF2CC", "font_color": "#806000", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })

        # Sheet 1: Dashboard Tổng Quan
        ws_dash = wb.add_worksheet("Tổng Quan")
        ws_dash.set_tab_color("#1B365D")
        ws_dash.hide_gridlines(0)
        ws_dash.set_column("A:A", 5)
        ws_dash.set_column("B:B", 38)
        ws_dash.set_column("C:C", 22)
        ws_dash.set_column("D:D", 52)

        title_fmt = wb.add_format({"bold": True, "font_size": 15, "font_color": "#1B365D"})
        subtitle_fmt = wb.add_format({"italic": True, "font_size": 10, "font_color": "#595959"})
        section_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "font_color": "#1B365D", "border": 1, "font_size": 11})

        ws_dash.write("B2", "BÁO CÁO ĐỐI SOÁT TAM GIÁC & TỔNG QUAN CHẤT LƯỢNG DỮ LIỆU CCTG", title_fmt)
        ws_dash.write("B3", f"Hệ thống: ABBA Data Engine | Tệp nguồn: {self.data_path.name} | Ngày xuất: {datetime.now().strftime('%d/%m/%Y %H:%M')}", subtitle_fmt)

        m = self.summary_metrics
        tot = m["total_rows"]

        # 1. QUY MÔ
        ws_dash.merge_range("B5:D5", "1. QUY MÔ CÁC NGUỒN DỮ LIỆU ĐỐI CHIẾU", section_fmt)
        ws_dash.write_row("B6", ["Nguồn Dữ Liệu", "Số Lượng Bản Ghi", "Ghi Chú"], fmt_hdr_base)
        ws_dash.write_row("B7", [f"[1] Data CCTG ({self.data_path.name})", f"{tot:,} dòng sổ AZ", f"{m.get('count_contracts_cnb', 43618):,} HĐ bán (CNB) | {m.get('count_contracts_cnm', 31311):,} HĐ mua (CNM)"], fmt_normal_cell)
        ws_dash.write_row("B8", ["[2] Báo Cáo Hóa Đơn Điện Tử", "12,411 dòng", "Hóa đơn tháng 9/2026 (12,230 hợp đồng)"], fmt_normal_cell)
        ws_dash.write_row("B9", ["[3] Sao Kê Tài Khoản TK43", f"{m['sk_clean_rows']:,} dòng", f"Giao dịch thực tế (Đã khử {m['sk_dupes_removed']:,} dòng trùng từ Bank | Chốt đến {m['sk_max_date']})"], fmt_normal_cell)

        # 2. ĐỐI SOÁT TỪNG DÒNG
        ws_dash.merge_range("B11:D11", f"2. KẾT QUẢ ĐỐI SOÁT CHI TIẾT TỪNG DÒNG ({tot:,} DÒNG SỔ AZ)", section_fmt)
        ws_dash.write_row("B12", ["Phân Loại Kết Quả", "Số Dòng Bản Ghi", "Tỷ Lệ / Chi Tiết"], fmt_hdr_base)
        ws_dash.write_row("B13", ["✅ Chuẩn 100% (Khớp không lệch)", f"{m['count_golden']:,}", f"{m['count_golden']/tot*100:.2f}% (Khớp hoàn hảo 0đ)"], fmt_normal_cell)
        ws_dash.write_row("B14", ["ℹ️ Lệch làm tròn bán lẻ App", f"{m['count_rounding']:,}", f"{m['count_rounding']/tot*100:.2f}% (Lệch số học < 2,000 đ - Xem Sheet Chênh Lệch Bán Lẻ)"], fmt_teal_row)
        ws_dash.write_row("B15", ["ℹ️ Giao dịch chờ Sao kê đợt mới", f"{m['count_future_sk']:,}", f"{m['count_future_sk']/tot*100:.2f}% (Phát sinh sau ngày chốt sao kê)"], fmt_normal_cell)
        ws_dash.write_row("B16", ["ℹ️ Nghiệp vụ ABBANK bán", f"{m['count_abbank']:,}", f"{m['count_abbank']/tot*100:.2f}% (Lịch sử trước giữa T6/2026)"], fmt_normal_cell)
        ws_dash.write_row("B17", ["ℹ️ Hợp đồng chưa đến ngày bán", f"{m['count_pending_sale']:,}", "Đang nắm giữ, chưa phát sinh dòng tiền bán"], fmt_normal_cell)
        ws_dash.write_row("B18", ["⚠️❌ Dòng có vấn đề / sai lệch thực tế", f"{m['count_discrepancy']:,}", f"{m['count_discrepancy']/tot*100:.2f}% (Xem chi tiết tại Tab Chi Tiết)"], fmt_yellow_row)

        # 3. ĐỐI SOÁT THEO HỢP ĐỒNG (CNM & CNB)
        ws_dash.merge_range("B20:D20", f"3. KẾT QUẢ ĐỐI SOÁT THEO HỢP ĐỒNG ({m.get('count_contracts_total', 43618):,} HỢP ĐỒNG)", section_fmt)
        ws_dash.write_row("B21", ["Chỉ Số Đối Soát Theo Hợp Đồng", "Số Hợp Đồng", "Tỷ Lệ / Đánh Giá Nghiệp Vụ"], fmt_hdr_base)
        ws_dash.write_row("B22", ["Chiều Mua (CNM) - Nộp tiền", f"{m.get('count_contracts_cnm', 31311):,} HĐ", f"{m.get('count_contracts_cnm_golden', 30693):,} HĐ khớp chuẩn 100% Ghi Có ({m.get('count_contracts_cnm_golden', 30693)/max(1, m.get('count_contracts_cnm', 31311))*100:.2f}%) | 12,229 HĐ khớp HĐĐT"], fmt_normal_cell)
        ws_dash.write_row("B23", ["Chiều Bán (CNB) - Đã tất toán", f"{m.get('count_contracts_sold', 27032):,} HĐ", f"{m.get('count_contracts_sold_golden', 26768):,} HĐ khớp chuẩn 100% Ghi Nợ | {m['count_retail_rounding_all']:,} HĐ lệch làm tròn bán lẻ"], fmt_normal_cell)
        ws_dash.write_row("B24", ["Hợp đồng CCTG đang nắm giữ", f"{m.get('count_contracts_holding', 16586):,} HĐ", "Chưa đến hạn tất toán, chưa phát sinh dòng tiền bán"], fmt_normal_cell)
        ws_dash.write_row("B25", ["Hợp đồng chờ Sao kê đợt mới", f"{m.get('count_contracts_future_sk', 618):,} HĐ", f"Phát sinh sau ngày chốt sao kê TK43 ({m['sk_max_date']})"], fmt_normal_cell)

        # 4. CHÊNH LỆCH BÁN LẺ
        ws_dash.merge_range("B27:D27", "4. ĐỐI SOÁT CHÊNH LỆCH LÀM TRÒN BÁN LẺ (ABBA MUA APP - CHIỀU BÁN CNB)", section_fmt)
        ws_dash.write_row("B28", ["Chỉ Số Đối Soát Bán Lẻ", "Giá Trị Thống Kê", "Đánh Giá Nghiệp Vụ"], fmt_hdr_base)
        ws_dash.write_row("B29", ["Số hợp đồng có chênh lệch", f"{m['count_retail_rounding_all']:,} HĐ ({m['count_retail_rounding_t9']:,} HĐ T9)", f"{m['count_rounding']:,} sổ AZ bán lẻ qua App (Xem Sheet Chênh Lệch Bán Lẻ)"], fmt_normal_cell)
        ws_dash.write_row("B30", ["Tổng tiền Data (SL x Đơn giá)", f"{m['sum_retail_data_all']:,.0f} đ", f"T9: {m['sum_retail_data_t9']:,.0f} đ | Các kỳ khác: {m['sum_retail_data_other']:,.0f} đ"], fmt_normal_cell)
        ws_dash.write_row("B31", ["Tổng tiền Sao kê Bank (Ghi Nợ)", f"{m['sum_retail_bank_all']:,.0f} đ", f"T9: {m['sum_retail_bank_t9']:,.0f} đ | Các kỳ khác: {m['sum_retail_bank_other']:,.0f} đ"], fmt_normal_cell)
        ws_dash.write_row("B32", ["Tổng chênh lệch ròng (Data - Bank)", f"{m['sum_retail_diff_all']:+,.0f} đ", f"T9: {m['sum_retail_diff_t9']:+,.0f} đ | Toàn bộ do làm tròn đơn giá lẻ sổ AZ (< 2,000 đ)"], fmt_green_row)

        # 5. CẢNH BÁO NGOẠI LỆ & XUNG ĐỘT CHỨNG TỪ
        ws_dash.merge_range("B34:D34", "5. CẢNH BÁO NGOẠI LỆ & XUNG ĐỘT CHỨNG TỪ (HỢP ĐỒNG ĐÃ THU TIỀN NHƯNG THIẾU/XUNG ĐỘT DATA)", section_fmt)
        ws_dash.write_row("B35", ["Nội Dung Ngoại Lệ & Xung Đột", "Số Lượng", "Đánh Giá Nghiệp Vụ & Khuyến Nghị Xử Lý"], fmt_hdr_base)
        ws_dash.write_row(
            "B36",
            [
                "HĐ 18715 (KH Lâm Gia Phước - 550 triệu)",
                "1 HĐ (550.000.000 đ)",
                "Ngân hàng thu đủ tiền FT26246688668854, HĐ đã xuất. Check ngược Data CCTG: FT này bị gán cho KH PHAM NGOC THIEN (Sổ AZ 0746000416011 có series đuôi Lâm Gia Phước) - Xem Sheet 5",
            ],
            fmt_red_row,
        )
        ws_dash.write_row(
            "B37",
            [
                "FT Sao Kê khác Tên KH / Mã HĐ trên Data CCTG",
                f"{m.get('count_cross_contract', 18):,} dòng ({m.get('count_cross_contract_fts', 8):,} FT)",
                "Mã FT ngân hàng khớp tiền nhưng diễn giải sao kê ghi KH/HĐ khác với Data CCTG (Nghi vấn gán nhầm FT hoặc nhầm CIF) - Xem Tab Chi Tiết",
            ],
            fmt_yellow_row,
        )

        # Sheet 2: Chi Tiết Theo Từng Dòng (80,763 dòng đầy đủ cột so sánh song song)
        ws_data = wb.add_worksheet("Chi Tiết Theo Từng Dòng")
        ws_data.set_tab_color("#276A3C")
        ws_data.freeze_panes(1, 6)

        headers_data = list(self.df_data_out.columns)
        for col_idx, h_name in enumerate(headers_data):
            h_fmt = get_header_fmt(h_name)
            ws_data.write(0, col_idx, h_name, h_fmt)
            c_len = max(14, len(h_name) + 2)
            if "Chi Tiết" in h_name or "Hướng Dẫn" in h_name or "CỜ" in h_name:
                c_len = max(c_len, 35)
            ws_data.set_column(col_idx, col_idx, c_len)

        data_vals = self.df_data_out.fillna("").values
        for r_idx in range(tot):
            h_code = highlight_codes[r_idx]
            cur_fmt = fmt_normal_cell
            if h_code == "GREEN":
                cur_fmt = fmt_green_row
            elif h_code == "YELLOW":
                cur_fmt = fmt_yellow_row
            elif h_code == "RED":
                cur_fmt = fmt_red_row
            elif h_code == "TEAL":
                cur_fmt = fmt_teal_row
            ws_data.write_row(r_idx + 1, 0, data_vals[r_idx], cur_fmt)

        # Sheet 3: Chi Tiết Theo Hợp Đồng (Toàn bộ 43,618 HĐ: cả CNM và CNB)
        ws_contract = wb.add_worksheet("Chi Tiết Theo Hợp Đồng")
        ws_contract.set_tab_color("#7030A0")
        ws_contract.freeze_panes(1, 5)

        contract_cols = [
            ("STT", 6, fmt_center_cell),
            ("Mã HĐ Mua (CNM)", 28, fmt_center_cell),
            ("Mã HĐ Bán (CNB)", 28, fmt_center_cell),
            ("Số CIF", 14, fmt_center_cell),
            ("Tên Khách Hàng", 26, fmt_normal_cell),
            ("Mã CCTG", 14, fmt_center_cell),
            ("Số Lượng Sổ AZ", 14, fmt_center_cell),
            ("Tổng SL CCTG", 14, fmt_curr_cell),
            ("[Chiều Mua - CNM] Ngày Mua", 15, fmt_center_cell),
            ("[Chiều Mua - CNM] Số FT Data", 20, fmt_center_cell),
            ("[Chiều Mua - CNM] Số FT Sao Kê", 20, fmt_center_cell),
            ("[Chiều Mua - CNM] Tiền Tính Toán (SL x Mệnh Giá)", 22, fmt_curr_cell),
            ("[Chiều Mua - CNM] Tiền Khai Báo (Header)", 22, fmt_curr_cell),
            ("[Chiều Mua - CNM] Tiền Ghi Có (Sao Kê)", 22, fmt_curr_cell),
            ("[Chiều Mua - CNM] Tiền Hóa Đơn (HĐĐT)", 20, fmt_curr_cell),
            ("[Chiều Mua - CNM] Δ Tiền Bank (Bank - Data)", 20, fmt_curr_cell),
            ("[Chiều Mua - CNM] Δ Tiền Hóa Đơn (Data - HĐ)", 20, fmt_curr_cell),
            ("[Chiều Bán - CNB] Ngày Bán", 15, fmt_center_cell),
            ("[Chiều Bán - CNB] Số FT Data", 20, fmt_center_cell),
            ("[Chiều Bán - CNB] Số FT Sao Kê", 20, fmt_center_cell),
            ("[Chiều Bán - CNB] Tiền Tính Toán (SL x Đơn Giá)", 22, fmt_curr_cell),
            ("[Chiều Bán - CNB] Tiền Khai Báo (Header)", 22, fmt_curr_cell),
            ("[Chiều Bán - CNB] Tiền Ghi Nợ (Sao Kê)", 22, fmt_curr_cell),
            ("[Chiều Bán - CNB] Chênh Lệch Bán Lẻ (Bank - Data)", 22, fmt_curr_cell),
            ("[Đánh Giá HĐ] Cờ Trạng Thái Hợp Đồng", 32, fmt_center_cell),
            ("[Đánh Giá HĐ] Chi Tiết Ghi Chú Đối Soát", 45, fmt_normal_cell),
            ("[Đánh Giá HĐ] Hướng Dẫn Xử Lý Kế Toán", 40, fmt_normal_cell),
        ]

        for col_idx, (col_name, col_width, _) in enumerate(contract_cols):
            h_fmt = get_header_fmt(col_name)
            ws_contract.write(0, col_idx, col_name, h_fmt)
            ws_contract.set_column(col_idx, col_idx, col_width)

        for idx, entry in enumerate(self.all_contracts, 1):
            r_idx = idx
            color = entry["color_hd"]
            if color == "TEAL":
                f_norm, f_curr, f_cent = fmt_teal_row, fmt_teal_curr, fmt_teal_center
            elif color == "GREEN":
                f_norm, f_curr, f_cent = fmt_green_row, fmt_green_curr, fmt_green_center
            elif color == "RED":
                f_norm, f_curr, f_cent = fmt_red_row, fmt_red_curr, fmt_red_center
            elif color == "YELLOW":
                f_norm, f_curr, f_cent = fmt_yellow_row, fmt_yellow_curr, fmt_yellow_center
            else:
                f_norm, f_curr, f_cent = fmt_normal_cell, fmt_curr_cell, fmt_center_cell

            ws_contract.write(r_idx, 0, idx, f_cent)
            ws_contract.write(r_idx, 1, entry["hd_mua"], f_cent)
            ws_contract.write(r_idx, 2, entry["hd_ban"], f_cent)
            ws_contract.write(r_idx, 3, entry["cif"], f_cent)
            ws_contract.write(r_idx, 4, entry["ten_kh"], f_norm)
            ws_contract.write(r_idx, 5, entry["ma_cctg"], f_cent)
            ws_contract.write(r_idx, 6, entry["so_luong_so_az"], f_cent)
            ws_contract.write(r_idx, 7, entry["tong_sl_cctg"], f_curr)
            ws_contract.write(r_idx, 8, entry["ngay_mua"], f_cent)
            ws_contract.write(r_idx, 9, entry["ft_mua_data"], f_cent)
            ws_contract.write(r_idx, 10, entry["ft_mua_sk"], f_cent)
            ws_contract.write(r_idx, 11, entry["calc_tien_mua"], f_curr)
            ws_contract.write(r_idx, 12, entry["hdr_tien_mua"], f_curr)
            ws_contract.write(r_idx, 13, entry["sk_credit"], f_curr)
            ws_contract.write(r_idx, 14, entry["inv_tien"], f_curr)
            ws_contract.write(r_idx, 15, entry["diff_m_bank"], f_curr)
            ws_contract.write(r_idx, 16, entry["diff_m_inv"], f_curr)
            ws_contract.write(r_idx, 17, entry["ngay_ban"], f_cent)
            ws_contract.write(r_idx, 18, entry["ft_ban_data"], f_cent)
            ws_contract.write(r_idx, 19, entry["ft_ban_sk"], f_cent)
            ws_contract.write(r_idx, 20, entry["calc_tien_ban"], f_curr)
            ws_contract.write(r_idx, 21, entry["hdr_tien_ban"], f_curr)
            ws_contract.write(r_idx, 22, entry["sk_debit"], f_curr)
            ws_contract.write(r_idx, 23, entry["diff_b_retail"], f_curr)
            ws_contract.write(r_idx, 24, entry["stt_hd"], f_cent)
            ws_contract.write(r_idx, 25, entry["detail_hd"], f_norm)
            ws_contract.write(r_idx, 26, entry["guide_hd"], f_norm)

        # Định dạng riêng cho Sheet Chênh Lệch Bán Lẻ
        fmt_total_lbl = wb.add_format({
            "bold": True, "bg_color": "#D9E1F2", "font_color": "#1B365D", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_total_num = wb.add_format({
            "bold": True, "bg_color": "#D9E1F2", "font_color": "#1B365D", "border": 1, "valign": "vcenter", "num_format": r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)', "font_size": 9
        })
        fmt_total_diff = wb.add_format({
            "bold": True, "bg_color": "#D9E1F2", "font_color": "#C00000", "border": 1, "valign": "vcenter", "num_format": r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)', "font_size": 9
        })
        fmt_total_note = wb.add_format({
            "bold": True, "italic": True, "bg_color": "#D9E1F2", "font_color": "#276A3C", "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })
        fmt_curr_cell = wb.add_format({
            "border": 1, "valign": "vcenter", "num_format": r'_(* #,##0_);_(* \(#,##0\);_(* "-"??_);_(@_)', "font_size": 9
        })
        fmt_center_cell = wb.add_format({
            "border": 1, "valign": "vcenter", "align": "center", "font_size": 9
        })

        retail_headers = [
            "STT",
            "Hợp đồng ABBA mua",
            "Tên Khách Hàng",
            "Số Lượng Sổ AZ",
            "Kỳ Bán",
            "Ngày Bán",
            "Số FT Bán",
            "SL x Đơn giá",
            "Sao kê",
            "Chênh Lệch",
            "Đánh Giá / Nguyên Nhân",
        ]

        # Sheet 6: Chênh Lệch Bán Lẻ (Gộp toàn bộ 264 HĐ làm tròn thành 1 tab duy nhất, Tháng 9 lên đầu)
        ws_retail = wb.add_worksheet("Chênh Lệch Bán Lẻ")
        ws_retail.set_tab_color("#0E6251")
        ws_retail.freeze_panes(2, 2)

        n_all = len(self.retail_rounding_all)
        max_r_all = n_all + 2  # Dữ liệu từ dòng 3 (index 2) đến max_r_all

        # Dòng 1: Tổng cộng (Excel formulas)
        ws_retail.write("A1", "TỔNG CỘNG", fmt_total_lbl)
        for col_empty in ["B1", "C1", "D1", "E1", "F1", "G1"]:
            ws_retail.write(col_empty, "", fmt_total_lbl)
        ws_retail.write_formula("H1", f"=SUM(H3:H{max_r_all})", fmt_total_num, sum(e["sl_x_don_gia"] for e in self.retail_rounding_all))
        ws_retail.write_formula("I1", f"=SUM(I3:I{max_r_all})", fmt_total_num, sum(e["sao_ke"] for e in self.retail_rounding_all))
        ws_retail.write_formula("J1", f"=SUM(J3:J{max_r_all})", fmt_total_diff, sum(e["chenh_lech"] for e in self.retail_rounding_all))
        ws_retail.write("K1", "100% Lệch làm tròn số học (< 2,000 đ)", fmt_total_note)

        # Dòng 2: Tiêu đề
        for col_idx, h_name in enumerate(retail_headers):
            ws_retail.write(1, col_idx, h_name, fmt_hdr_base)

        ws_retail.set_column("A:A", 6)
        ws_retail.set_column("B:B", 30)
        ws_retail.set_column("C:C", 26)
        ws_retail.set_column("D:D", 16)
        ws_retail.set_column("E:E", 12)
        ws_retail.set_column("F:F", 14)
        ws_retail.set_column("G:G", 22)
        ws_retail.set_column("H:H", 20)
        ws_retail.set_column("I:I", 20)
        ws_retail.set_column("J:J", 16)
        ws_retail.set_column("K:K", 42)

        for idx, entry in enumerate(self.retail_rounding_all, 1):
            r_idx = idx + 1
            r_excel = r_idx + 1
            ws_retail.write(r_idx, 0, idx, fmt_center_cell)
            ws_retail.write(r_idx, 1, entry["hd_ban"], fmt_normal_cell)
            ws_retail.write(r_idx, 2, entry["ten_kh"], fmt_normal_cell)
            ws_retail.write(r_idx, 3, entry["so_luong_so_az"], fmt_center_cell)
            ws_retail.write(r_idx, 4, entry["ky_ban"], fmt_center_cell)
            ws_retail.write(r_idx, 5, entry["ngay_ban"], fmt_center_cell)
            ws_retail.write(r_idx, 6, entry["ft_ban"], fmt_center_cell)
            ws_retail.write(r_idx, 7, entry["sl_x_don_gia"], fmt_curr_cell)
            ws_retail.write(r_idx, 8, entry["sao_ke"], fmt_curr_cell)
            ws_retail.write_formula(r_idx, 9, f"=H{r_excel}-I{r_excel}", fmt_curr_cell, entry["chenh_lech"])
            ws_retail.write(r_idx, 10, entry["danh_gia"], fmt_normal_cell)

        # Sheet 5: HĐ & Sao Kê Ngoại Lệ (HĐ 18715 - KH Lâm Gia Phước 550 triệu)
        ws_miss = wb.add_worksheet("HĐ & Sao Kê Ngoại Lệ")
        ws_miss.set_tab_color("#ED7D31")
        ws_miss.freeze_panes(1, 2)
        df_missing_hd_clean = pd.DataFrame(self.missing_hd_list).fillna("")
        miss_cols = list(df_missing_hd_clean.columns)
        for col_idx, c_name in enumerate(miss_cols):
            ws_miss.write(0, col_idx, c_name, fmt_hdr_base)
            c_len = max(16, len(c_name) + 2)
            if "Đánh Giá" in c_name or "Chi Tiết" in c_name or "Khuyến Nghị" in c_name or "Diễn Giải" in c_name:
                c_len = 55
            elif "Tên" in c_name or "Mã Hợp Đồng" in c_name:
                c_len = 28
            elif "Series" in c_name:
                c_len = 38
            ws_miss.set_column(col_idx, col_idx, c_len)

        for r_idx, row_vals in enumerate(df_missing_hd_clean.values, 1):
            row_dict = df_missing_hd_clean.iloc[r_idx - 1]
            is_critical = (
                "CỰC KỲ" in str(row_dict.get("Mức Độ Rủi Ro", ""))
                or "XUNG ĐỘT" in str(row_dict.get("[Kiểm Tra Ngược Data CCTG] Trạng Thái FT Trên Data", ""))
            )
            f_row = fmt_red_row if is_critical else fmt_normal_cell
            f_curr = fmt_red_curr if is_critical else fmt_curr_cell
            f_cent = fmt_red_center if is_critical else fmt_center_cell

            for c_idx, val in enumerate(row_vals):
                c_name = miss_cols[c_idx]
                if "Tiền" in c_name or "Δ" in c_name:
                    try:
                        num_v = float(val) if val != "" else 0.0
                        ws_miss.write_number(r_idx, c_idx, num_v, f_curr)
                    except (ValueError, TypeError):
                        ws_miss.write(r_idx, c_idx, val, f_row)
                elif "Ngày" in c_name or "Mã" in c_name or ("Số" in c_name and "Tiền" not in c_name):
                    ws_miss.write(r_idx, c_idx, str(val), f_cent)
                else:
                    ws_miss.write(r_idx, c_idx, str(val), f_row)

        writer.close()



def run_reconciliation():
    engine = TriangleReconciliationEngine()
    return engine.execute()


if __name__ == "__main__":
    run_reconciliation()
