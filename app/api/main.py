"""
FastAPI Server Entry Point for Reconciliation System
===================================================
"""

from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from app.api.routes import router
from app.api.service import ReconciliationService


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Preload reconciliation data on startup
    print("\n[FastAPI Startup] Tải trước bộ đệm kết quả đối soát vào bộ nhớ...")
    service = ReconciliationService.get_instance()
    try:
        summary = service.get_summary()
        print(
            f"[FastAPI Startup] Sẵn sàng phục vụ: Tổng {summary['total_rows']:,} dòng. "
            f"Chuẩn tuyệt đối: {summary['count_golden']:,} ({summary['pct_golden']}%).\n"
        )
    except Exception as e:
        print(f"[FastAPI Startup] Lưu ý khởi tạo: {e}\n")
    yield


app = FastAPI(
    title="CCTG Triangle Bank Reconciliation & Data API",
    description="Backend REST API for cross-reconciling Bank Statements (TK43), Invoices, and CCTG Contract Data.",
    version="2.0.0",
    lifespan=lifespan,
)

# Enable CORS for frontend clients
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(router)

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("app.api.main:app", host="0.0.0.0", port=8000, reload=True)
