# CCTG Bank & Triangle Reconciliation REST API (v2.0)

Backend REST API phục vụ đối soát đa chiều tam giác giữa **Data CCTG** ↔ **Báo Cáo Hóa Đơn** ↔ **Sao Kê Ngân Hàng TK43** và chuẩn hóa dữ liệu.

---

## 🚀 1. Cách Khởi Động API Server

### Cách 1: Click chạy trực tiếp file Batch (Windows)
Double-click vào file:
👉 `3_Khoi_Chay_API_Server.bat`

### Cách 2: Dòng lệnh terminal
```bash
python -m uvicorn app.api.main:app --host 0.0.0.0 --port 8000 --reload
```

Sau khi khởi chạy:
- **API Base URL:** `http://localhost:8000`
- **Tài liệu trực quan Swagger UI:** [http://localhost:8000/docs](http://localhost:8000/docs)
- **Tài liệu ReDoc:** [http://localhost:8000/redoc](http://localhost:8000/redoc)

---

## 📌 2. Danh Sách Các Endpoints Cốt Lõi

| Phương thức | Đường dẫn | Chức năng & Tham số |
| :---: | :--- | :--- |
| `GET` | `/api/health` | Kiểm tra trạng thái hoạt động (Health check v2.0.0). |
| `POST` | `/api/reconcile/run` | Kích hoạt chạy đối soát từ file mặc định hoặc đường dẫn tùy chọn (`ReconcileRunRequest`). |
| `GET` | `/api/reconcile/summary` | Lấy dữ liệu tổng quan tỷ lệ đạt (`pct_golden`, `count_golden`, `count_discrepancy`, `sk_max_date`...). |
| `GET` | `/api/reconcile/problems` | Lấy danh sách các dòng có vấn đề / sai lệch. Có phân trang (`page`, `page_size`) và tìm kiếm (`search`). |
| `GET` | `/api/reconcile/export` | Tải xuống file Excel báo cáo đối soát 4 sheet hoàn chỉnh (`Data_abba_Chuan_Va_DoiSoat.xlsx`). |

---

## 💻 3. Hướng Dẫn Tích Hợp Frontend (React / TypeScript)

### 3.1. Gọi API Lấy Thống Kê Tổng Quan (Summary)
```typescript
export async function fetchReconciliationSummary() {
  const res = await fetch("http://localhost:8000/api/reconcile/summary");
  if (!res.ok) throw new Error("Failed to load summary");
  return await res.json();
}
```

### 3.2. Gọi API Lấy Danh Sách Chi Tiết Dòng Có Vấn Đề (Phân Trang & Tìm Kiếm)
```typescript
export async function fetchProblemRecords(
  page: number = 1,
  pageSize: number = 50,
  search?: string
) {
  const params = new URLSearchParams({
    page: page.toString(),
    page_size: pageSize.toString(),
  });
  if (search) params.append("search", search);

  const res = await fetch(`http://localhost:8000/api/reconcile/problems?${params}`);
  if (!res.ok) throw new Error("Failed to load problem records");
  return await res.json();
}
```

### 3.3. Tải File Báo Cáo Excel 4 Sheet
```typescript
export function downloadExcelReport() {
  window.open("http://localhost:8000/api/reconcile/export", "_blank");
}
```

---

## 📂 4. Cấu Trúc Mã Nguồn Module API

```
app/api/
├── __init__.py
├── main.py          # Khởi tạo FastAPI app, lifespan cache, cấu hình CORS
├── routes.py        # Các API endpoints xử lý RESTful request/response
├── schemas.py       # Pydantic data models cho type validation
└── service.py       # Tầng Service điều phối engine và in-memory cache
```
