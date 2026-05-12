"""공통 설정 — .env에서 값을 읽어 전체 모듈에 제공.

.env 탐색 우선순위
  1. 환경변수 HWPMINE_ENV 가 지정한 경로
  2. 현재 작업 디렉터리의 ./.env
  3. %APPDATA%\\hwpmine\\.env   (Windows 사용자 설정)
  4. ~/.config/hwpmine/.env     (기타 플랫폼)
"""

import os
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    raise SystemExit("python-dotenv 필요: pip install python-dotenv")


def _candidate_env_paths() -> list[Path]:
    paths: list[Path] = []

    override = os.environ.get("HWPMINE_ENV")
    if override:
        paths.append(Path(override).expanduser())

    paths.append(Path.cwd() / ".env")

    if os.name == "nt":
        appdata = os.environ.get("APPDATA")
        if appdata:
            paths.append(Path(appdata) / "hwpmine" / ".env")
    else:
        xdg = os.environ.get("XDG_CONFIG_HOME")
        base = Path(xdg) if xdg else Path.home() / ".config"
        paths.append(base / "hwpmine" / ".env")

    return paths


def _load_env() -> Path | None:
    for p in _candidate_env_paths():
        if p.is_file():
            load_dotenv(p, override=False)
            return p
    return None


ENV_PATH = _load_env()

# ── DB ────────────────────────────────────────────────────────
DB_HOST     = os.getenv("DB_HOST", "127.0.0.1")
DB_PORT     = int(os.getenv("DB_PORT", "3306"))
DB_USER     = os.getenv("DB_USER", "root")
DB_PASSWORD = os.getenv("DB_PASSWORD", "")
DB_NAME     = os.getenv("DB_NAME", "hwp_documents")
# 적재 테이블은 HWP/PDF 가 공용 — 두 파이프라인 모두 같은 테이블에 행을 쌓고,
# 검색은 extension 컬럼으로 구분하거나 통합 검색한다.
DB_TABLE    = os.getenv("DB_TABLE", "documents")

# ── 스캔 ──────────────────────────────────────────────────────
SCAN_DRIVES = [d.strip() for d in os.getenv("SCAN_DRIVES", r"C:\,D:\\").split(",") if d.strip()]

# ── CSV ───────────────────────────────────────────────────────
# 스캔 단계는 HWP/PDF 가 별도 CSV — 각자의 파서로 적재하기 위함.
CSV_FILE     = os.getenv("CSV_FILE",     "hwp_file_list.csv")
PDF_CSV_FILE = os.getenv("PDF_CSV_FILE", "pdf_file_list.csv")

# ── 적재 튜닝 ─────────────────────────────────────────────────
COMMIT_EVERY  = int(os.getenv("COMMIT_EVERY",  "50"))
COM_RESTART   = int(os.getenv("COM_RESTART",   "500"))
PARSE_TIMEOUT = int(os.getenv("PARSE_TIMEOUT", "60"))

# PDF 본문 추출은 CPU-바운드 — 멀티프로세스로 병렬 처리.
# 기본은 실행 머신의 논리 CPU 수에 맞춰 동적으로 결정.
# 환경변수 PDF_WORKERS / CLI --workers 로 override 가능 (1 이상 정수).
def _detect_pdf_workers() -> int:
    env = os.getenv("PDF_WORKERS", "0").strip()
    try:
        n = int(env)
    except ValueError:
        n = 0
    if n > 0:
        return n
    return max(1, os.cpu_count() or 1)


PDF_WORKERS = _detect_pdf_workers()


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
