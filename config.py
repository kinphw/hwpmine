"""공통 설정 — .env에서 값을 읽어 전체 모듈에 제공."""

import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    raise SystemExit("python-dotenv 필요: pip install python-dotenv")

load_dotenv(Path(__file__).parent / ".env")

# ── DB ────────────────────────────────────────────────────────
DB_HOST     = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT     = int(os.getenv("DB_PORT", "3306"))
DB_USER     = os.getenv("DB_USER", "root")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")
DB_NAME     = os.getenv("DB_NAME", "hwp_documents")
DB_TABLE    = os.getenv("DB_TABLE", "documents")

# ── 스캔 ──────────────────────────────────────────────────────
SCAN_DRIVES = [d.strip() for d in os.getenv("SCAN_DRIVES", r"C:\,D:\\").split(",") if d.strip()]

# ── CSV ───────────────────────────────────────────────────────
CSV_FILE = os.getenv("CSV_FILE", "hwp_file_list.csv")

# ── 적재 튜닝 ─────────────────────────────────────────────────
COMMIT_EVERY  = int(os.getenv("COMMIT_EVERY",  "50"))
COM_RESTART   = int(os.getenv("COM_RESTART",   "500"))
PARSE_TIMEOUT = int(os.getenv("PARSE_TIMEOUT", "60"))


def get_db_config(use_db: bool = True) -> dict:
    cfg = {
        "host":    DB_HOST,
        "port":    DB_PORT,
        "user":    DB_USER,
        "password": DB_PASSWORD,
        "charset": "utf8mb4",
    }
    if use_db:
        cfg["database"] = DB_NAME
    return cfg
