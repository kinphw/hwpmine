"""공통 설정 — .env에서 값을 읽어 전체 모듈에 제공.

.env 탐색 우선순위
  1. 환경변수 DOCMINE_ENV (또는 구 HWPMINE_ENV) 가 지정한 경로
  2. 현재 작업 디렉터리의 ./.env
  3. %APPDATA%\\docmine\\.env   (Windows 사용자 설정)
  4. ~/.config/docmine/.env     (기타 플랫폼)
  5. (3·4 폴백) %APPDATA%\\hwpmine\\.env / ~/.config/hwpmine/.env

구 hwpmine 경로 또는 HWPMINE_ENV 환경변수에서 로드된 경우 stderr 에 1회
안내를 출력하며, 다음 마이너 버전에서 호환이 제거된다.

구 사용자 .env 호환 처리
- PDF_CSV_FILE 항목이 없으면 CSV_FILE 옆에 같은 폴더/확장자로 자동 배치
  (예: hwp_file_list.csd → pdf_file_list.csd).
- DOCMINE_ENV / PDF_WORKERS 등 새 항목들은 모두 sensible default 보유.
"""

import os
import sys
from pathlib import Path

try:
    from dotenv import load_dotenv
except ImportError:
    raise SystemExit("python-dotenv 필요: pip install python-dotenv")


def _exe_folder_env() -> Path | None:
    """PyInstaller 로 묶인 단일 exe 로 실행 중일 때, 그 exe 와 같은 폴더의
    .env 를 후보로 반환. cmd 의 CWD 와 Explorer 더블클릭 시 CWD 가
    다를 수 있어, 사용자가 'exe 옆에 .env 두면 된다' 는 직관을 보존."""
    if getattr(sys, "frozen", False):
        try:
            return Path(sys.executable).resolve().parent / ".env"
        except Exception:
            return None
    return None


def _candidate_env_paths() -> list[tuple[Path, bool]]:
    """[(path, is_legacy), …] 우선순위 순.

    is_legacy=True 는 구 hwpmine 경로/환경변수에서 유래한 후보. 로드 성공
    시 사용자에게 1회 안내를 출력하고, 다음 마이너 버전에서 제거된다.
    """
    out: list[tuple[Path, bool]] = []

    docmine_env = os.environ.get("DOCMINE_ENV")
    hwpmine_env = os.environ.get("HWPMINE_ENV")
    if docmine_env:
        out.append((Path(docmine_env).expanduser(), False))
    elif hwpmine_env:
        out.append((Path(hwpmine_env).expanduser(), True))

    # PyInstaller 단일 exe 모드: exe 옆의 .env. Explorer 더블클릭 시
    # CWD 가 항상 exe 폴더는 아닐 수 있어 명시 후보로 둠.
    exe_env = _exe_folder_env()
    if exe_env is not None:
        out.append((exe_env, False))

    out.append((Path.cwd() / ".env", False))

    if os.name == "nt":
        appdata = os.environ.get("APPDATA")
        if appdata:
            out.append((Path(appdata) / "docmine" / ".env", False))
            out.append((Path(appdata) / "hwpmine" / ".env", True))
    else:
        xdg = os.environ.get("XDG_CONFIG_HOME")
        base = Path(xdg) if xdg else Path.home() / ".config"
        out.append((base / "docmine" / ".env", False))
        out.append((base / "hwpmine" / ".env", True))

    return out


def _legacy_notice(p: Path, env_var: bool) -> None:
    where = "HWPMINE_ENV 환경변수" if env_var else f"구 hwpmine 경로 ({p})"
    sys.stderr.write(
        f"[docmine] {where} 에서 .env 를 로드했습니다.\n"
        f"[docmine] 다음 마이너 버전부터는 docmine 전용 경로만 인식합니다.\n"
        f"[docmine]   권장: {p} 를 "
        f"{Path(os.environ.get('APPDATA', '%APPDATA%')) / 'docmine' / '.env'} 로 이동\n"
    )


def _load_env() -> Path | None:
    legacy_env_var = (
        not os.environ.get("DOCMINE_ENV") and bool(os.environ.get("HWPMINE_ENV"))
    )
    for p, is_legacy in _candidate_env_paths():
        if p.is_file():
            load_dotenv(p, override=False)
            if is_legacy:
                _legacy_notice(p, env_var=legacy_env_var)
            return p
    # 후보 어느 곳에도 .env 가 없으면 default 가 적용되는데, 운영망에서
    # 이게 묵묵히 동작하면 DB 인증 시점에 의문의 OperationalError 가 난다.
    # (특히 root 계정이 auth_gssapi_client 로 설정된 MariaDB 에서.)
    # 사용자가 즉시 알 수 있도록 후보 경로와 함께 경고.
    try:
        candidates = "\n    ".join(str(p) for p, _ in _candidate_env_paths())
        sys.stderr.write(
            "[docmine] 경고: .env 파일을 찾지 못했습니다. 기본값(DB_USER=root,"
            " DB_PASSWORD='') 으로 동작합니다.\n"
            "  탐색한 위치:\n"
            f"    {candidates}\n"
            "  .env.example 을 복사해 위 경로 중 한 곳에 .env 로 두고 값을 채우세요.\n"
        )
        sys.stderr.flush()
    except Exception:
        pass
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
CSV_FILE = os.getenv("CSV_FILE", "hwp_file_list.csv")


def _default_pdf_csv() -> str:
    """PDF_CSV_FILE 기본값 — env 미지정 시 CSV_FILE 와 동일 폴더·확장자로 배치.

    구 hwpmine .env 에는 PDF_CSV_FILE 이 없으므로, HWP CSV 경로를 단서로
    PDF CSV 도 같은 위치에 자동 배치되도록 한다. 'hwp' 토큰이 stem 에
    있으면 'pdf' 로 치환하고, 없으면 stem 앞에 'pdf_' 를 붙인다.

    예) hwp_file_list.csd → pdf_file_list.csd
        D:/data/hwp_files.csv → D:/data/pdf_files.csv
        D:/data/my_list.csv   → D:/data/pdf_my_list.csv
    """
    explicit = os.getenv("PDF_CSV_FILE")
    if explicit:
        return explicit
    hwp = Path(CSV_FILE)
    stem = hwp.stem
    if "hwp" in stem:
        new_stem = stem.replace("hwp", "pdf")
    elif "HWP" in stem:
        new_stem = stem.replace("HWP", "PDF")
    else:
        new_stem = "pdf_" + stem
    return str(hwp.with_name(new_stem + hwp.suffix))


PDF_CSV_FILE = _default_pdf_csv()

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
