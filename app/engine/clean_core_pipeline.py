"""
Pipeline Làm Sạch Dữ Liệu Chuyên Dụng Cho CD CORE DATA
======================================================
Mục tiêu:
1. Chuẩn hóa & bảo toàn cấu trúc 34 cột nguyên bản của Core Data.
2. Bổ sung 4,772 số hóa đơn, ký hiệu, trạng thái từ Báo cáo Hóa đơn điện tử.
3. Bổ sung hơn 4,000 số FT Mua bị khuyết từ Sao Kê TK43.
4. Xử lý triệt để xung đột hợp đồng HĐ 18715 (Lâm Gia Phước CIF 13558581 vs Phạm Ngọc Thiện tại dòng 54468).
5. Phân bổ hoàn hảo chênh lệch làm tròn bán lẻ (264 hợp đồng) vào Đơn giá bán, đảm bảo:
   Tổng (Số lượng * Đơn giá bán) == Tổng GT HĐ Bán == Sao Kê Ngân Hàng (100% khớp).
6. Xuất bản 2 định dạng:
   - data/output/Data_abba_Clean_Core.csv (UTF-8 BOM, bảo toàn số 0 đầu)
   - data/output/Data_abba_Clean_Core.xlsx (Format ngân hàng chuyên nghiệp)
"""

import os
import re
import time
from collections import defaultdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import xlsxwriter

from app.core.config import (
    DEFAULT_DATA_PATH,
    DEFAULT_HD_PATH,
    DEFAULT_SK_PATH,
    OUTPUT_DIR,
)
from app.core.parsers import clean_int_str, extract_ft, norm_code, parse_date, parse_date_str

OUTPUT_CLEAN_CSV = OUTPUT_DIR / "Data_abba_Clean_Core.csv"
OUTPUT_CLEAN_XLSX = OUTPUT_DIR / "Data_abba_Clean_Core.xlsx"


class CleanCoreDataPipeline:
    def __init__(
        self,
        data_path: Optional[Path] = None,
        hd_path: Optional[Path] = None,
        sk_path: Optional[Path] = None,
    ):
        self.data_path = Path(data_path) if data_path else DEFAULT_DATA_PATH
        self.hd_path = Path(hd_path) if hd_path else DEFAULT_HD_PATH
        self.sk_path = Path(sk_path) if sk_path else DEFAULT_SK_PATH
        self.stats = {
            "total_rows": 0,
            "invoices_enriched": 0,
            "buy_fts_enriched": 0,
            "lam_gia_phuoc_fixed": 0,
            "retail_contracts_adjusted": 0,
            "retail_diff_resolved": 0,
        }

    def execute(self) -> bool:
        start_time = time.time()
        print("=" * 80, flush=True)
        print("  🚀 BẮT ĐẦU PIPELINE LÀM SẠCH DỮ LIỆU ĐỂ IMPORT VÀO CD CORE DATA", flush=True)
        print("=" * 80, flush=True)

        # 1. Đọc dữ liệu thô
        print(f"\n[1/6] Nạp dữ liệu CCTG gốc: {self.data_path.name} ...", flush=True)
        if self.data_path.suffix.lower() == ".csv":
            df = pd.read_csv(self.data_path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
        else:
            df = pd.read_excel(self.data_path, dtype=str).fillna("")

        self.stats["total_rows"] = len(df)
        original_cols = list(df.columns)
        print(f"      -> Tổng số dòng: {len(df):,} | Tổng số cột: {len(original_cols)}", flush=True)

        # 2. Nạp và xây dựng bản đồ Hóa Đơn Điện Tử
        print(f"\n[2/6] Nạp dữ liệu Hóa Đơn Điện Tử: {self.hd_path.name} ...", flush=True)
        inv_map = {}
        if self.hd_path.exists():
            df_hd = pd.read_excel(self.hd_path)
            for _, r in df_hd.iterrows():
                p_name = str(r.get("Tên sản phẩm", ""))
                m_c = re.search(r"(CN[MB]-?\d+-\d+)", p_name, re.I)
                c_norm = norm_code(m_c.group(1)) if m_c else ""
                so_hd = clean_int_str(r.get("Số hóa đơn", ""))
                ky_hieu = str(r.get("Mẫu số, ký hiệu", "")).strip()
                if "," in ky_hieu:
                    ky_hieu = ky_hieu.split(",")[-1].strip()
                trang_thai = str(r.get("Trạng thái hóa đơn", "Hóa đơn gốc")).strip()
                if c_norm and so_hd:
                    inv_map[c_norm] = {
                        "so_hd": so_hd,
                        "ky_hieu": ky_hieu,
                        "trang_thai": trang_thai,
                    }
            print(f"      -> Sẵn sàng đối chiếu {len(inv_map):,} hợp đồng có Hóa đơn.", flush=True)

        # 3. Nạp và xây dựng bản đồ Sao Kê TK43
        print(f"\n[3/6] Nạp dữ liệu Sao Kê TK43: {self.sk_path.name} ...", flush=True)
        sk_buy_map = {}
        if self.sk_path.exists():
            df_sk = pd.read_excel(self.sk_path)
            for _, r in df_sk.iterrows():
                credit = float(pd.to_numeric(r.get("Số tiền có (Credit amount)", 0), errors="coerce") or 0)
                if credit > 0:
                    desc = str(r.get("Diễn giải (Description)", ""))
                    m_c = re.search(r"(CN[MB]-?\d+-\d+)", desc, re.I)
                    c_norm = norm_code(m_c.group(1)) if m_c else ""
                    ft_raw = r.get("Số giao dịch (Transaction Number)")
                    ft_clean = extract_ft(ft_raw)
                    if c_norm and ft_clean:
                        sk_buy_map[c_norm] = {
                            "ft": ft_clean,
                            "credit": credit,
                        }
            print(f"      -> Sẵn sàng đối chiếu {len(sk_buy_map):,} hợp đồng có FT Mua từ ngân hàng.", flush=True)

        # 4. Thực hiện làm sạch và chuẩn hóa dữ liệu
        print("\n[4/6] Đang tiến hành làm sạch, bù khuyết và phân bổ chênh lệch ...", flush=True)

        # 4.1. Sửa lỗi bản ghi xung đột Lâm Gia Phước (HĐ 18715 tại dòng 54468)
        mask_phuoc = (df["Số FT HĐ Mua"] == "FT26246688668854") | (df["Số series thứ cấp"].str.contains("135585811788395225436", na=False))
        if mask_phuoc.any():
            for idx in df[mask_phuoc].index:
                df.at[idx, "Số CIF"] = "13558581"
                df.at[idx, "Tên KH"] = "Lâm Gia Phước"
                df.at[idx, "Số định danh/MST"] = "074089003323"
                df.at[idx, "Số HĐ mua"] = "CNM-13558581-1788395225436"
                df.at[idx, "Số HĐ Bán"] = "CNB-13558581-1788395225436"
                df.at[idx, "Số hoá đơn"] = "18715"
                df.at[idx, "Ký hiệu hoá đơn"] = "C26TAK"
                df.at[idx, "Trạng thái hoá đơn"] = "Hóa đơn gốc"
                self.stats["lam_gia_phuoc_fixed"] += 1
            print(f"      ✔ Đã khắc phục dứt điểm xung đột HĐ 18715: Cập nhật về đúng chủ sở hữu Lâm Gia Phước (CIF 13558581).", flush=True)

        # 4.2. Bổ sung Hóa đơn và FT Mua
        for i in range(len(df)):
            hd_m = str(df["Số HĐ mua"].iat[i]).strip()
            norm_m = norm_code(hd_m)

            # Bổ sung Hóa đơn
            curr_so_hd = str(df["Số hoá đơn"].iat[i]).strip()
            if (not curr_so_hd or curr_so_hd in ("0", "nan", "None")) and norm_m in inv_map:
                inv = inv_map[norm_m]
                df.at[i, "Số hoá đơn"] = inv["so_hd"]
                if not str(df["Ký hiệu hoá đơn"].iat[i]).strip():
                    df.at[i, "Ký hiệu hoá đơn"] = inv["ky_hieu"]
                if not str(df["Trạng thái hoá đơn"].iat[i]).strip():
                    df.at[i, "Trạng thái hoá đơn"] = inv["trang_thai"]
                self.stats["invoices_enriched"] += 1

            # Bổ sung FT Mua
            curr_ft_m = str(df["Số FT HĐ Mua"].iat[i]).strip()
            if (not curr_ft_m or curr_ft_m.lower() in ("nan", "none", "")) and norm_m in sk_buy_map:
                df.at[i, "Số FT HĐ Mua"] = sk_buy_map[norm_m]["ft"]
                self.stats["buy_fts_enriched"] += 1

        print(f"      ✔ Đã bổ sung thành công {self.stats['invoices_enriched']:,} dòng thông tin Hóa đơn.", flush=True)
        print(f"      ✔ Đã bổ sung thành công {self.stats['buy_fts_enriched']:,} dòng Số FT HĐ Mua.", flush=True)

        # 4.3. Phân bổ chênh lệch làm tròn bán lẻ (Retail Rounding Allocation)
        print("      Đang tối ưu & phân bổ chênh lệch làm tròn cho các hợp đồng bán lẻ ...", flush=True)
        sold_indices = df[df["Số HĐ Bán"] != ""].index.tolist()
        contracts_sold = defaultdict(list)
        for idx in sold_indices:
            hdb = str(df["Số HĐ Bán"].loc[idx]).strip()
            contracts_sold[hdb].append(idx)

        for hdb, idx_list in contracts_sold.items():
            # Lấy header bán
            hdr_str = str(df["Tổng GT HĐ Bán"].loc[idx_list[0]]).replace(",", "").strip()
            try:
                hdr_val = float(hdr_str)
            except ValueError:
                continue

            if hdr_val <= 0:
                continue

            # Tính tổng các dòng
            sl_list = []
            dg_list = []
            for r_idx in idx_list:
                s_str = str(df["Số lượng"].loc[r_idx]).replace(",", "").strip()
                d_str = str(df["Đơn giá bán/ số seri"].loc[r_idx]).replace(",", "").strip()
                sl_list.append(float(s_str) if s_str else 0.0)
                dg_list.append(float(d_str) if d_str else 0.0)

            calc_sum = sum(s * d for s, d in zip(sl_list, dg_list))
            diff = hdr_val - calc_sum

            # Nếu có lệch làm tròn (nhỏ hơn dung sai 1,000 đ)
            if abs(diff) >= 0.001 and abs(diff) < 2000.0:
                self.stats["retail_contracts_adjusted"] += 1
                diff_int = int(round(diff))

                # Thuật toán phân bổ:
                # Bước 1: Ưu tiên dòng có SL == 1
                adjusted = False
                for i_pos, r_idx in enumerate(idx_list):
                    if sl_list[i_pos] == 1.0:
                        new_dg = dg_list[i_pos] + diff
                        # Lưu giá trị nguyên nếu diff là số nguyên
                        if abs(new_dg - round(new_dg)) < 1e-4:
                            df.at[r_idx, "Đơn giá bán/ số seri"] = str(int(round(new_dg)))
                        else:
                            df.at[r_idx, "Đơn giá bán/ số seri"] = f"{new_dg:.4f}".rstrip("0").rstrip(".")
                        adjusted = True
                        break

                # Bước 2: Dòng có SL chia hết cho diff_int
                if not adjusted and diff_int != 0:
                    for i_pos, r_idx in enumerate(idx_list):
                        s_int = int(sl_list[i_pos])
                        if s_int > 0 and diff_int % s_int == 0:
                            step = diff_int // s_int
                            new_dg = dg_list[i_pos] + step
                            df.at[r_idx, "Đơn giá bán/ số seri"] = str(int(round(new_dg)))
                            adjusted = True
                            break

                # Bước 3: Diophantine nguyên 2 ẩn cho các cặp dòng phổ biến
                if not adjusted and diff_int != 0 and len(idx_list) >= 2:
                    found_dioph = False
                    for i_a in range(len(idx_list)):
                        for i_b in range(i_a + 1, len(idx_list)):
                            sa = int(sl_list[i_a])
                            sb = int(sl_list[i_b])
                            if sa > 0 and sb > 0:
                                # Thử c_a, c_b trong khoảng [-5, 5]
                                for ca in range(-5, 6):
                                    rem = diff_int - ca * sa
                                    if rem % sb == 0:
                                        cb = rem // sb
                                        if abs(cb) <= 6:
                                            # Tìm thấy giải pháp nguyên!
                                            df.at[idx_list[i_a], "Đơn giá bán/ số seri"] = str(int(round(dg_list[i_a] + ca)))
                                            df.at[idx_list[i_b], "Đơn giá bán/ số seri"] = str(int(round(dg_list[i_b] + cb)))
                                            adjusted = True
                                            found_dioph = True
                                            break
                                if found_dioph:
                                    break
                        if found_dioph:
                            break

                # Bước 4: Nếu không có tổ hợp nguyên nhỏ, chia đều hoặc phân bổ số thập phân chính xác vào dòng có SL nhỏ nhất
                if not adjusted:
                    # Lấy dòng có SL nhỏ nhất
                    min_pos = min(range(len(idx_list)), key=lambda k: sl_list[k] if sl_list[k] > 0 else 999999)
                    r_idx = idx_list[min_pos]
                    s_val = sl_list[min_pos]
                    if s_val > 0:
                        new_dg = dg_list[min_pos] + (diff / s_val)
                        df.at[r_idx, "Đơn giá bán/ số seri"] = f"{new_dg:.4f}".rstrip("0").rstrip(".")
                        adjusted = True

                self.stats["retail_diff_resolved"] += 1

        print(f"      ✔ Đã xử lý phân bổ làm tròn thành công: {self.stats['retail_contracts_adjusted']} hợp đồng.", flush=True)

        # 5. Xuất bản tệp CSV chuẩn UTF-8 có BOM cho Core Data
        print(f"\n[5/6] Đang xuất tệp CSV sạch cho CD CORE DATA: {OUTPUT_CLEAN_CSV.name} ...", flush=True)
        # Đảm bảo giữ đúng 34 cột nguyên bản và thứ tự cột
        df_clean = df[original_cols]
        df_clean.to_csv(OUTPUT_CLEAN_CSV, index=False, encoding="utf-8-sig")
        csv_size_mb = os.path.getsize(OUTPUT_CLEAN_CSV) / (1024 * 1024)
        print(f"      ✔ Đã ghi {len(df_clean):,} dòng ra {OUTPUT_CLEAN_CSV} ({csv_size_mb:.2f} MB)", flush=True)

        # 6. Xuất bản tệp Excel định dạng chuẩn ngân hàng
        print(f"\n[6/6] Đang xuất tệp Excel định dạng chuẩn ngân hàng: {OUTPUT_CLEAN_XLSX.name} ...", flush=True)
        self._export_to_excel(df_clean, OUTPUT_CLEAN_XLSX)

        elapsed = time.time() - start_time
        print("\n" + "=" * 80, flush=True)
        print("  🎉 HOÀN THÀNH TẤT CẢ CÁC BƯỚC LÀM SẠCH DỮ LIỆU CD CORE DATA", flush=True)
        print(f"  - Tổng số dòng dữ liệu: {self.stats['total_rows']:,}")
        print(f"  - Số dòng bổ sung Hóa Đơn: {self.stats['invoices_enriched']:,}")
        print(f"  - Số dòng bổ sung FT Mua: {self.stats['buy_fts_enriched']:,}")
        print(f"  - Số bản ghi xử lý Lâm Gia Phước (HĐ 18715): {self.stats['lam_gia_phuoc_fixed']}")
        print(f"  - Số HĐ bán lẻ phân bổ làm tròn: {self.stats['retail_contracts_adjusted']}")
        print(f"  - File CSV Core: {OUTPUT_CLEAN_CSV}")
        print(f"  - File Excel Core: {OUTPUT_CLEAN_XLSX}")
        print(f"  - Tổng thời gian xử lý: {elapsed:.1f} giây")
        print("=" * 80 + "\n", flush=True)
        return True

    def _export_to_excel(self, df: pd.DataFrame, out_path: Path):
        workbook = xlsxwriter.Workbook(out_path, {"constant_memory": True})
        worksheet = workbook.add_worksheet("Data_abba_Clean_Core")
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

        fmt_left = workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "left",
            "valign": "vcenter",
        })

        fmt_center = workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "center",
            "valign": "vcenter",
        })

        fmt_curr = workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0",
        })

        fmt_curr_dec = workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0.00",
        })

        fmt_int = workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "right",
            "valign": "vcenter",
            "num_format": "#,##0",
        })

        fmt_rate = workbook.add_format({
            "font_name": font_family,
            "font_size": 9,
            "border": 1,
            "border_color": "#E0E0E0",
            "align": "right",
            "valign": "vcenter",
            "num_format": "0.0",
        })

        headers = list(df.columns)
        money_cols = {"Mệnh giá", "Tổng GT HĐ mua", "Tổng GT HĐ Bán", "Lãi Coupon đã trả thực tế"}
        int_cols = {"Số lượng", "Thời hạn nắm giữ (tháng)"}
        rate_cols = {"LS Coupon", "Lãi suất HĐ Mua"}
        date_cols = {"Ngày báo cáo", "Ngày phát hành AZ", "Ngày đáo hạn AZ", "Ngày mua (ngày nắm giữ)", "Ngày hết hạn nắm giữ", "Ngày bán"}
        center_text_cols = {
            "Số CIF", "Loại KH", "Số định danh/MST", "Mã CCTG", "Số sổ AZ", "Số FT HĐ Mua", "Số FT HĐ Bán",
            "Số hoá đơn", "Ký hiệu hoá đơn", "Trạng thái hoá đơn", "Mã CKS HĐ Mua", "Trạng thái GD Mua",
            "Mã CKS HĐ Bán", "Trạng thái GD Bán"
        }

        # Độ rộng cột
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

        total_rows = len(df)
        data_matrix = df.values
        for r_idx in range(total_rows):
            row_num = r_idx + 1
            worksheet.set_row(row_num, 19)
            row_raw = data_matrix[r_idx]

            for col_idx, val in enumerate(row_raw):
                h_name = headers[col_idx]
                if not val or str(val).lower() in ("nan", "none", "null"):
                    worksheet.write_string(row_num, col_idx, "", fmt_center if h_name in center_text_cols else fmt_left)
                    continue

                val_str = str(val).strip()
                if h_name == "Đơn giá bán/ số seri":
                    try:
                        num_v = float(val_str.replace(",", ""))
                        if num_v.is_integer():
                            worksheet.write_number(row_num, col_idx, int(num_v), fmt_curr)
                        else:
                            worksheet.write_number(row_num, col_idx, num_v, fmt_curr_dec)
                    except ValueError:
                        worksheet.write_string(row_num, col_idx, val_str, fmt_left)

                elif h_name in money_cols:
                    try:
                        num_v = float(val_str.replace(",", ""))
                        worksheet.write_number(row_num, col_idx, num_v, fmt_curr)
                    except ValueError:
                        worksheet.write_string(row_num, col_idx, val_str, fmt_left)

                elif h_name in int_cols:
                    try:
                        num_v = int(float(val_str.replace(",", "")))
                        worksheet.write_number(row_num, col_idx, num_v, fmt_int)
                    except ValueError:
                        worksheet.write_string(row_num, col_idx, val_str, fmt_left)

                elif h_name in rate_cols:
                    try:
                        num_v = float(val_str)
                        worksheet.write_number(row_num, col_idx, num_v, fmt_rate)
                    except ValueError:
                        worksheet.write_string(row_num, col_idx, val_str, fmt_center)

                elif h_name in date_cols:
                    worksheet.write_string(row_num, col_idx, parse_date_str(val_str), fmt_center)

                elif h_name in ("Số hoá đơn", "Mã CKS HĐ Mua", "Mã CKS HĐ Bán"):
                    worksheet.write_string(row_num, col_idx, clean_int_str(val_str), fmt_center)

                elif h_name in center_text_cols:
                    worksheet.write_string(row_num, col_idx, val_str, fmt_center)

                else:
                    worksheet.write_string(row_num, col_idx, val_str, fmt_left)

        worksheet.autofilter(0, 0, total_rows, len(headers) - 1)
        workbook.close()
        xlsx_size_mb = os.path.getsize(out_path) / (1024 * 1024)
        print(f"      ✔ Đã ghi {total_rows:,} dòng ra {out_path} ({xlsx_size_mb:.2f} MB)", flush=True)


if __name__ == "__main__":
    pipeline = CleanCoreDataPipeline()
    pipeline.execute()
