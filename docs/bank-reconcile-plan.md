# Kế Hoạch Triển Khai: Hệ Thống Đối Soát Dữ Liệu Ngân Hàng & Sao Kê Tài Khoản
**Mã Kế Hoạch:** `bank-reconcile-plan.md`  
**Mục tiêu:** Đối chiếu tự động, đa chiều giữa dữ liệu Bank gửi (CCTG / Nghiệp vụ ABBA) và Sao kê tài khoản (TK43) dựa trên 4 tiêu chí cốt lõi: Mã FT – Số tiền – Mã Hợp Đồng / Sổ AZ – Ngày hiệu lực.  
**Ngày tạo:** 2026-09-29  
**Người phụ trách / Agents:** `project-planner`, `backend-specialist`, `database-architect`, `qa-automation-engineer`

---

## 📌 1. Bối Cảnh & Mục Tiêu Nghiệp Vụ

### 1.1. Dữ liệu đầu vào
- **Tệp 1 (Sao kê TK43):** File Excel `data/saoke_TK43.xlsx` chứa giao dịch thực tế phát sinh tại tài khoản ngân hàng (Số giao dịch FT, Ngày hiệu lực, Diễn giải, Chiều Debit/Credit, Số tiền).
- **Tệp 2 (Dữ liệu Bank gửi / CCTG ABBA):** File CSV `data/Data_abba_20260929.csv` chứa thông tin hợp đồng chuyển nhượng tiền gửi (Mã CCTG, Số HĐ mua/bán, Số sổ AZ, Mệnh giá, Số lượng, Tổng GT, Ngày mua/bán, Số FT HĐ mua/bán).

### 1.2. Mục tiêu đối soát 4 tiêu chí cốt lõi
1. **Mã FT:** Khớp nối định danh giao dịch ngân hàng giữa 2 nguồn dữ liệu.
2. **Số tiền (Debit/Amount):** Dung sai sai số tuyệt đối `< 1.0 VND` (xử lý giao dịch gộp 1 FT ứng với nhiều hợp đồng).
3. **Mã Hợp Đồng (CNM/CNB) hoặc Số sổ AZ:** Bóc tách chuỗi linh hoạt từ nội dung diễn giải tự do (loại trừ dấu nối `-`, khoảng trắng).
4. **Ngày hiệu lực (Value Date):** So khớp ngày hạch toán với ngày giao dịch, bao gồm dung sai chuyển tiếp cuối tuần (Thứ 7 / Chủ Nhật → Thứ Hai).

---

## 🏗️ 2. Kiến Trúc Bộ Quy Tắc (Reconciliation Rules Matrix)

Hệ thống triển khai **10 quy tắc nghiệp vụ chuẩn** được phân bổ thành 2 nhóm:

### 🏛️ Nhóm A: Đối Soát Chéo Bank ↔ Nghiệp Vụ (BANK-01 ➜ BANK-07)
| Mã Rule | Tên Quy Tắc | Mức Độ | Điều Kiện Thỏa Mãn & Hành Động |
| :--- | :--- | :---: | :--- |
| **BANK-01** | **Khớp Chuẩn 1:1 Toàn Diện** | `CRITICAL` | Đồng thời thỏa mãn 4 tiêu chí: Khớp FT 1:1, Số tiền trong dung sai, Chiều Ghi Nợ (Debit), Ngày giá trị (gồm rollover cuối tuần), Số HĐ/AZ có trong diễn giải. |
| **BANK-02** | **Không Tìm Thấy FT** | `HIGH` | Mã FT trên sao kê ngân hàng không tồn tại trong dữ liệu nghiệp vụ (Bank đã trừ tiền nhưng nghiệp vụ thiếu hồ sơ). |
| **BANK-03** | **Sai Hợp Đồng / Sổ AZ** | `MEDIUM` | Khớp FT nhưng Số HĐ / Sổ AZ trong diễn giải không khớp với thông tin lưu trên hệ thống CCTG. |
| **BANK-04** | **Lệch Số Tiền Ghi Nợ** | `HIGH` | Khớp FT nhưng số tiền Debit trên sao kê lệch so với Tổng GT hợp đồng CCTG. |
| **BANK-05** | **Lệch Ngày Giá Trị** | `LOW` | Khớp FT nhưng lệch ngày ngoài khung dung sai chuyển tiếp ngày nghỉ cuối tuần. |
| **BANK-06** | **Không Phải Chiều Ghi Nợ** | `CRITICAL` | Khớp FT mua CCTG nhưng chiều giao dịch trên sao kê lại là Ghi Có (Credit). |
| **BANK-07** | **Giao Dịch Trùng / Gộp** | `HIGH` | 1 mã FT ngân hàng ánh xạ cho nhiều dòng bản ghi hợp đồng CCTG (Giao dịch gộp). |

### 🔍 Nhóm B: Kiểm Soát Chất Lượng Dữ Liệu Nghiệp Vụ (DQ-01 ➜ DQ-03)
| Mã Rule | Tên Quy Tắc | Mức Độ | Điều Kiện Thỏa Mãn & Hành Động |
| :--- | :--- | :---: | :--- |
| **DQ-01** | **Trùng Series Thứ Cấp** | `HIGH` | Phát hiện số series thứ cấp bị trùng lặp trên các dòng hồ sơ CCTG. |
| **DQ-02** | **Thiếu Series (Gán Mã Tạm)** | `LOW` | Bản ghi CCTG thiếu số series thứ cấp, tự động gán mã định danh tạm thời (`a1`, `a2`,...). |
| **DQ-03** | **Lệch Mệnh Giá Cấp Hợp Đồng** | `CRITICAL` | Kiểm tra tổng mệnh giá theo Hợp đồng: $\text{ABS}\left(\text{Tổng GT HĐ} - \sum (\text{Mệnh giá} \times \text{Số lượng})\right) < 0.01$. |

---

## 📋 3. Lộ Trình Triển Khai Chi Tiết (Phases & Task Breakdown)

```
┌────────────────────────────────────────────────────────┐
│ Giai đoạn 1: Chuẩn Hóa Engine & Tiền Xử Lý Dữ Liệu    │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│ Giai đoạn 2: Tối Ưu Hóa Bộ Phân Tích & Đối Soát 4 Tiêu Chí
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│ Giai đoạn 3: Tích Hợp API / Stored Procedure & Backend │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│ Giai đoạn 4: Báo Cáo Excel Đa Sheet & Giao Diện Web UI │
└──────────────────────────┬─────────────────────────────┘
                           │
                           ▼
┌────────────────────────────────────────────────────────┐
│ Giai đoạn 5: Kiểm Thử Tự Động & Bàn Giao Vận Hành      │
└────────────────────────────────────────────────────────┘
```

### Giai đoạn 1: Chuẩn Hóa Engine & Tiền Xử Lý Dữ Liệu
- [x] **Task 1.1:** Chuẩn hóa module đọc đa định dạng ngày:
  - Bank: `DD/MM/YYYY` hoặc `Timestamp`.
  - CCTG Mua: Chuỗi số 8 ký tự `YYYYMMDD`.
  - CCTG Bán: Định dạng `M/D/YYYY H:M:S AM/PM`.
- [x] **Task 1.2:** Xây dựng bảng tra cứu Hash Map nhanh (FT Lookup Index) liên kết 1 FT sang tập hợp các bản ghi CCTG Mua/Bán (`O(1)` access).
- [x] **Task 1.3:** Xử lý làm sạch chuỗi hợp đồng (`norm_code`): Loại bỏ ký tự đặc biệt, gạch ngang, khoảng trắng, đưa về chữ in hoa chuẩn.

### Giai đoạn 2: Tối Ưu Hóa Bộ Phân Tích & Đối Soát 4 Tiêu Chí
- [x] **Task 2.1:** Hoàn thiện bộ quy tắc dung sai ngày (`dates_match`): Xử lý tự động chuyển tiếp cuối tuần (Thứ 7 / Chủ Nhật → Thứ Hai).
- [x] **Task 2.2:** Logic gom nhóm giao dịch 1:N (`BANK-07`): Tự động tính tổng giá trị các HĐ con để đối chiếu với món nợ Debit trên sao kê.
- [x] **Task 2.3:** Bóc tách hợp đồng & sổ AZ trong nội dung diễn giải: Tìm kiếm chuỗi con sau khi chuẩn hóa, hỗ trợ cả hợp đồng mua (`CNM`), hợp đồng bán (`CNB`) và sổ AZ.
- [x] **Task 2.4:** Sửa thuật toán `DQ-03` chuyển từ cấp Dòng sang cấp Hợp đồng (`_hd_ref`), loại bỏ 100% cảnh báo giả (false positives).

### Giai đoạn 3: Tích Hợp Cơ Sở Dữ Liệu & Backend
- [ ] **Task 3.1:** Đồng bộ cấu trúc logic Python sang SQL Stored Procedure (`silver.usp_reconcile_12_rules`) nếu chạy đối soát trực tiếp trong Data Warehouse.
- [ ] **Task 3.2:** Xây dựng REST API endpoint (FastAPI / Express) phục vụ trigger tiến trình đối soát và nhận tiến độ (polling/SSE progress).
- [ ] **Task 3.3:** Lưu trữ vết lịch sử các phiên đối soát (Reconciliation Audit Trail / Run History).

### Giai đoạn 4: Báo Cáo Excel Đa Sheet & Giao Diện Web Dashboard
- [x] **Task 4.1:** Xuất bản tự động workbook Excel 12 sheet với bảng màu ngân hàng chuyên nghiệp (`Summary`, `Statistics`, `BANK-01` đến `BANK-07`, `DQ-01` đến `DQ-03`).
- [ ] **Task 4.2:** Kết nối dữ liệu đối soát lên giao diện Web `ReconciliationBankTab.tsx` trong hệ thống Portal để kiểm toán viên thao tác duyệt, lọc giao dịch ngoại lệ.
- [ ] **Task 4.3:** Cung cấp tính năng tải xuống báo cáo ngoại lệ dành riêng cho đội kế toán xử lý với ngân hàng.

### Giai đoạn 5: Kiểm Thử Tự Động & Bàn Giao Vận Hành
- [x] **Task 5.1:** Kiểm thử hồi quy với tập dữ liệu thực tế:
  - Sao kê TK43: **62,480 dòng**.
  - CCTG: **77,855 dòng**.
  - Tỷ lệ khớp nối chuẩn `BANK-01`: **22,224 dòng**.
  - Thời gian xử lý toàn trình: `< 60 giây`.
- [ ] **Task 5.2:** Xây dựng bộ test case giả lập biên (edge cases):
  - Lệch lẻ vài hào / vài đồng.
  - Diễn giải chứa nhiều số hợp đồng cùng lúc.
  - Giao dịch chuyển khoản qua các kỳ nghỉ lễ dài ngày (30/4, Tết).
- [ ] **Task 5.3:** Viết tài liệu hướng dẫn vận hành và cấu hình tham số dung sai (`reconcile_config.json`).

---

## 🎯 4. Tiêu Chí Nghiệm Thu (Acceptance Criteria)

1. **Tính chính xác:**
   - `BANK-01` đạt 100% độ tin cậy, không có bản ghi nào bị gán sai khi thiếu một trong 4 tiêu chí.
   - Không xuất hiện cảnh báo giả ở quy tắc mệnh giá `DQ-03`.
2. **Hiệu năng:**
   - Xử lý trên 100,000 dòng dữ liệu trong thời gian dưới 90 giây trên môi trường máy trạm chuẩn.
3. **Đầu ra hoàn chỉnh:**
   - File Excel có đầy đủ định dạng số tiền tệ `#,,##0`, tiêu đề rõ ràng, không bị lỗi gãy công thức hay font chữ.
   - Thống kê phân loại rõ ràng 4 mức độ: `CRITICAL`, `HIGH`, `MEDIUM`, `LOW`.

---

## 🚀 5. Hướng Dẫn Thực Thi Tiếp Theo (Next Steps)

1. Xem lại nội dung kế hoạch tại file: **[bank-reconcile-plan.md](file:///D:/D%E1%BB%B0%20%C3%81N%20ABBA/%C4%90%E1%BB%91i%20So%C3%A1t%20Raw%20APP/bank-reconcile-plan.md)**.
2. Để bắt đầu triển khai các bước backend hoặc giao diện web tiếp theo, bạn có thể chạy lệnh `/create` hoặc yêu cầu tôi thực hiện từng giai đoạn cụ thể.
