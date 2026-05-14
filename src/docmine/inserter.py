"""
Step 2 - HWP 배치 파싱 -> MariaDB
==================================
CSV 목록을 읽어 HWP/HWPX 본문을 파싱한 뒤 MariaDB에 적재.
워커 프로세스 격리로 COM 크래시 시 자동 재시작.

단독 실행:
  python inserter.py --create-db
  python inserter.py
  python inserter.py --start 0 --end 1000
  python inserter.py --csv other_list.csv
"""

import csv
import os
import re
import sys
import time
import html as _html
import argparse
import subprocess
import threading
import multiprocessing as mp
from pathlib import Path
from queue import Empty
from typing import Optional

from .hwp_parser import configure_logging

try:
    import pymysql
except ImportError:
    raise SystemExit("pymysql 필요: pip install pymysql")

from . import config

ERROR_LOG = "hwp_parse_errors.csv"


# ================================================================
# DDL / SQL
# ================================================================

DDL_DB = (
    f"CREATE DATABASE IF NOT EXISTS `{config.DB_NAME}`"
    " CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
)

DDL_TABLE = f"""
CREATE TABLE IF NOT EXISTS `{config.DB_TABLE}` (
    id           INT AUTO_INCREMENT PRIMARY KEY,
    directory    VARCHAR(1000)  NOT NULL,
    filename     VARCHAR(500)   NOT NULL,
    extension    VARCHAR(10)    NOT NULL,
    file_size    BIGINT         DEFAULT 0,
    file_mtime   VARCHAR(30),
    body_text    LONGTEXT,
    parse_status ENUM('success','error','skip','empty') DEFAULT 'success',
    error_msg    VARCHAR(1000),
    parsed_at    DATETIME       DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY   uq_file (directory(500), filename(255)),
    INDEX        idx_parse_status (parse_status),
    INDEX        idx_extension    (extension),
    INDEX        idx_filename     (filename(191))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
"""

# 기존 환경(이전 버전 ENUM) 보강용 — 스캔본 PDF 를 'empty' 로 표기 가능하도록.
ALTER_STATUS_ENUM = (
    f"ALTER TABLE `{config.DB_TABLE}` "
    "MODIFY COLUMN parse_status "
    "ENUM('success','error','skip','empty') DEFAULT 'success'"
)

# 기존 테이블(인덱스 미보유) 보강용. 이미 존재하면 try/except 로 무시.
#
# 인덱스가 가속하는 것 / 못 하는 것
#   ○ parse_status / extension : 메타 필터링
#   ○ filename(191)            : 파일명 prefix 검색·정렬
#   × LIKE '%kw%' (선행 %)     : 어떤 B-tree 인덱스도 도움 안 됨 — 풀스캔
#                                불가피. 본문 검색 가속은 search_gui 의
#                                쿼리를 MATCH AGAINST 로 바꾸고 FULLTEXT
#                                인덱스를 추가해야 가능(현재 미적용 —
#                                한국어 토크나이저 제약 + 본문 LONGTEXT
#                                FULLTEXT 빌드가 서버 다운 위험).
DDL_INDEXES = [
    f"CREATE INDEX idx_parse_status ON `{config.DB_TABLE}` (parse_status)",
    f"CREATE INDEX idx_extension    ON `{config.DB_TABLE}` (extension)",
    f"CREATE INDEX idx_filename     ON `{config.DB_TABLE}` (filename(191))",
]

INSERT_SQL = f"""
INSERT INTO `{config.DB_TABLE}`
    (directory, filename, extension, file_size, file_mtime,
     body_text, parse_status, error_msg)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
ON DUPLICATE KEY UPDATE
    body_text=VALUES(body_text), parse_status=VALUES(parse_status),
    error_msg=VALUES(error_msg), parsed_at=CURRENT_TIMESTAMP
"""


def _diagnose_db_error(err: Exception) -> str:
    """자주 보는 DB 연결 에러에 .env 상태·해결 힌트를 덧붙여 GUI 로그에 보여줌."""
    msg = str(err)
    lines: list[str] = []

    env_path = getattr(config, "ENV_PATH", None)
    if env_path is None:
        lines.append(
            f"  .env 미로드 - 기본값 사용 중: DB_USER='{config.DB_USER}', "
            f"DB_HOST='{config.DB_HOST}'"
        )
        lines.append(
            "  .env.example 을 복사해 다음 중 한 곳에 .env 로 두고 값을 채우세요:"
        )
        lines.append("    - 실행 폴더 (python.exe 와 같은 폴더)")
        lines.append("    - %APPDATA%\\docmine\\.env")
    else:
        lines.append(
            f"  .env 로드됨: {env_path}   DB_USER='{config.DB_USER}',"
            f" DB_HOST='{config.DB_HOST}'"
        )

    if "auth_gssapi_client" in msg:
        lines.append(
            "  → 'auth_gssapi_client' 는 Kerberos/AD 통합 인증 플러그인입니다."
        )
        lines.append(
            "    PyMySQL 은 미지원 → 해당 DB 사용자가 mysql_native_password 로"
            " 설정된 계정인지 확인하세요."
        )
    elif "Access denied" in msg:
        lines.append("  → 사용자/비밀번호를 확인하세요.")
    elif "Unknown database" in msg:
        lines.append(
            f"  → DB '{config.DB_NAME}' 가 서버에 없습니다. 서버에서 CREATE DATABASE 필요."
        )
    return "\n".join(lines)


def get_conn(use_db: bool = True):
    try:
        return pymysql.connect(**config.get_db_config(use_db))
    except pymysql.err.OperationalError as e:
        hint = _diagnose_db_error(e)
        if hint:
            # 원본 에러는 보존하면서 진단 메시지를 메시지 끝에 덧붙여
            # GUI 의 로그 패널에 멀티라인으로 표시되도록.
            args = list(e.args)
            if len(args) >= 2:
                args[1] = f"{args[1]}\n{hint}"
            else:
                args = [getattr(e, 'errno', 0), f"{e}\n{hint}"]
            raise pymysql.err.OperationalError(*args) from e
        raise


def create_db():
    conn = get_conn(use_db=False)
    with conn.cursor() as cur:
        cur.execute(DDL_DB)
        conn.select_db(config.DB_NAME)
        cur.execute(DDL_TABLE)
        # 이전 스키마에서 만든 테이블이라면 ENUM 에 'empty' 가 없을 수 있어
        # 보강. 이미 'empty' 가 포함돼 있으면 MariaDB 는 메타데이터-only 로 처리.
        try:
            cur.execute(ALTER_STATUS_ENUM)
        except Exception:
            pass
        # 기존 테이블에 인덱스가 없으면 추가. 이미 있으면 1061 (Duplicate
        # key name) 가 떨어지는데 무시.
        for ddl in DDL_INDEXES:
            try:
                cur.execute(ddl)
            except Exception:
                pass
    conn.commit()
    conn.close()
    print(f"  v DB [{config.DB_NAME}] / 테이블 [{config.DB_TABLE}] 준비 완료")


# ================================================================
# Windows Job Object — 부모 프로세스가 어떻게 죽든 자식까지 일괄 살처분.
# ================================================================
#
# mp.Pool/mp.Process 워커는 Python 의 daemon=True 라도 실제 종료는 부모의
# atexit 훅이 terminate() 를 호출하는 방식이라, GUI 강제종료/X클릭/
# Task Manager 같은 비정상 종료에서는 워커가 고아화되어 CPU·메모리를
# 계속 잡는다.
#
# Job Object 에 JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE 를 걸어두면 부모가
# 사라지는 순간 커널이 그 job 의 모든 프로세스를 즉시 종료한다. spawn 으로
# 만들어진 워커들은 별다른 옵션 없이 부모의 job 을 상속한다.
#
# HWP 워커가 COM 으로 띄우는 Hwp.exe 는 DCOM 활성화 경로에 따라 job 을
# 벗어날 수 있으므로 _kill_hwp() 와 시작 시점 청소가 보조 안전망이다.
_job_handle = None  # GC 되면 효과가 사라지므로 모듈 변수로 보관.


def _setup_kill_on_close_job() -> None:
    """현재 프로세스를 KILL_ON_JOB_CLOSE Job 에 할당. 이미 다른 job 에 속해
    있어 실패하면(예: 디버거/샌드박스 환경) 경고만 찍고 진행."""
    global _job_handle
    if os.name != "nt" or _job_handle is not None:
        return
    try:
        import win32api
        import win32con
        import win32job
    except ImportError:
        return
    try:
        job = win32job.CreateJobObject(None, "")
        info = win32job.QueryInformationJobObject(
            job, win32job.JobObjectExtendedLimitInformation
        )
        info["BasicLimitInformation"]["LimitFlags"] |= (
            win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
        )
        win32job.SetInformationJobObject(
            job, win32job.JobObjectExtendedLimitInformation, info
        )
        h = win32api.OpenProcess(
            win32con.PROCESS_ALL_ACCESS, False, win32api.GetCurrentProcessId()
        )
        win32job.AssignProcessToJobObject(job, h)
        _job_handle = job
    except Exception as e:
        print(f"  ⚠ Job Object 설정 실패(무시): {e}")


# ================================================================
# 진행바
# ================================================================

class PB:
    def __init__(self, total, offset=0):
        self.total = total
        self.offset = offset
        self.cur = self.ok = self.err = self.crash = self.skip = self.empty = 0
        self.t0 = time.time()

    def tick(self, status="success"):
        self.cur += 1
        if status == "success":
            self.ok += 1
        elif status == "skip":
            self.skip += 1
        elif status == "empty":
            # 본문 텍스트가 비어 있는 PDF (스캔본/이미지) — 에러와 분리해 표기.
            self.empty += 1
        elif status == "crash":
            self.crash += 1
        else:
            self.err += 1

        el  = time.time() - self.t0
        eta = el / self.cur * (self.total - self.cur) if self.cur else 0
        pct = self.cur / self.total
        w   = 30
        bar = "#" * int(w * pct) + "." * (w - int(w * pct))
        idx = self.offset + self.cur
        empty_str = f" empty:{self.empty}" if self.empty else ""
        skip_str = f" skip:{self.skip}" if self.skip else ""
        crash_str = f" crash:{self.crash}" if self.crash else ""
        print(
            f"\r  [{bar}] {idx}/{self.offset + self.total}"
            f"  ok:{self.ok} err:{self.err}{empty_str}{skip_str}{crash_str}"
            f"  ETA {int(eta//60)}:{int(eta%60):02d}  ",
            end="", flush=True,
        )

    def done(self):
        el = time.time() - self.t0
        empty_str = f" empty:{self.empty}" if self.empty else ""
        skip_str = f" skip:{self.skip}" if self.skip else ""
        crash_str = f" crash:{self.crash}" if self.crash else ""
        print(
            f"\n  완료: {int(el//60)}분{int(el%60)}초"
            f"  ok:{self.ok} err:{self.err}{empty_str}{skip_str}{crash_str}"
        )


# ================================================================
# 워커 프로세스
# ================================================================

def _kill_hwp():
    for p in ["Hwp.exe", "HwpFrame.exe"]:
        try:
            subprocess.run(["taskkill", "/F", "/IM", p],
                           capture_output=True, timeout=5)
        except Exception:
            pass


def _clean(text: str) -> str:
    text = text.replace(chr(0x02), "")
    text = text.replace(chr(0x05), "")
    text = text.replace(chr(0x0B), "\n")
    text = text.replace(chr(0x1C), "")
    text = text.replace(chr(0x1D), "")
    text = text.replace(chr(0x1E), "")
    text = text.replace(chr(0x1F), "")
    text = text.replace(chr(0xA0), " ")
    return re.sub(r"[\x00-\x08\x0c\x0e-\x1b]", "", text).strip()


def worker_main(task_q, result_q, kill_hwp=True):
    """
    별도 프로세스로 실행. COM 인스턴스를 유지하며 파일을 파싱.
    크래시 시 이 프로세스만 종료 -> 메인이 새 워커를 띄움.

    kill_hwp=False 로 호출하면 COM 재활용/종료 시 Hwp.exe taskkill 을 생략한다.
    extractor 처럼 사용자가 외부에서 한/글을 띄워두고 사용하는 경로에서 쓰임.
    """
    from .hwp_parser import ZipDocReader, SectionParser, HWPXDrmError

    zip_reader = ZipDocReader()
    sec_parser = SectionParser()
    com = None
    com_count = 0

    def _get_com():
        nonlocal com
        if com is None:
            import win32com.client as win32
            # DispatchEx 로 신규 한/글 프로세스를 강제 생성한다.
            # gencache.EnsureDispatch 는 ROT(Running Object Table)에서 기존
            # 한/글 인스턴스를 찾으면 거기에 어태치해버려, 사용자가 띄워둔
            # 한/글이 워커의 COM 정리 시 같이 종료되는 부작용이 있었음.
            com = win32.DispatchEx("HwpFrame.HwpObject")
            try:
                com.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            except Exception:
                pass
            com.XHwpWindows.Item(0).Visible = False
            threading.Thread(target=_popup_loop, daemon=True).start()
        return com

    def _popup_loop():
        try:
            import win32gui
        except ImportError:
            return
        BTNS = ["접근 허용(&A)", "접근 허용", "확인(&O)", "확인", "OK",
                "아니오(&N)", "예(&Y)", "취소(&C)", "취소",
                "저장(&Y)", "저장"]
        while True:
            try:
                def _on(h, _):
                    if not win32gui.IsWindowVisible(h):
                        return
                    def _c(c, _):
                        try:
                            if win32gui.GetClassName(c) != "Button":
                                return
                            t = win32gui.GetWindowText(c)
                            if any(b in t for b in BTNS):
                                win32gui.SendMessage(c, 0xF5, 0, 0)
                        except Exception:
                            pass
                    try:
                        win32gui.EnumChildWindows(h, _c, None)
                    except Exception:
                        pass
                win32gui.EnumWindows(_on, None)
            except Exception:
                pass
            time.sleep(0.3)

    def _com_extract(filepath: str) -> str:
        nonlocal com, com_count
        hwp = _get_com()
        try:
            hwp.SetMessageBoxMode(0x10000)
        except Exception:
            pass
        # DispatchEx(late-binding)는 optional 인자 디폴트가 안 채워지므로 3-arg 명시.
        # "forceopen:true" 로 손상/경고 파일도 강제로 열도록.
        hwp.Open(str(Path(filepath).absolute()), "", "forceopen:true")
        text = ""
        try:
            raw = hwp.GetTextFile("TEXT", "")
            if raw:
                text = _clean(_html.unescape(raw))
        except Exception:
            pass
        try:
            hwp.XHwpDocuments.Item(0).SetModified(False)
        except Exception:
            pass
        try:
            hwp.Run("FileClose")
        except Exception:
            pass
        try:
            hwp.SetMessageBoxMode(0xF0000)
        except Exception:
            pass
        com_count += 1
        if com_count % config.COM_RESTART == 0:
            try:
                del com
            except Exception:
                pass
            com = None
            if kill_hwp:
                _kill_hwp()
            time.sleep(1)
        return text

    while True:
        try:
            msg = task_q.get(timeout=5)
        except Empty:
            continue

        if msg is None:
            break

        idx, filepath, ext = msg
        try:
            if ext == ".hwpx":
                try:
                    doc  = zip_reader.read_document(Path(filepath), sec_parser)
                    text = doc.extract_text(skip_empty=True)
                except HWPXDrmError:
                    text = _com_extract(filepath)
            else:
                text = _com_extract(filepath)
            result_q.put((idx, "success", text, None))
        except Exception as e:
            result_q.put((idx, "error", None, str(e)[:900]))

    if com:
        try:
            del com
        except Exception:
            pass
    if kill_hwp:
        _kill_hwp()


# ================================================================
# 메인 프로세스
# ================================================================

def _spawn_worker(task_q, result_q):
    w = mp.Process(target=worker_main, args=(task_q, result_q), daemon=True)
    w.start()
    return w


def _load_existing_keys(conn, rows, chunk_size: int = 500):
    keys = {
        (row["directory"], row["filename"])
        for row in rows
        if row.get("directory") and row.get("filename")
    }
    if not keys:
        return set()

    existing = set()
    key_list = list(keys)
    for i in range(0, len(key_list), chunk_size):
        chunk = key_list[i:i + chunk_size]
        placeholders = ", ".join(["(%s, %s)"] * len(chunk))
        sql = (
            f"SELECT directory, filename FROM `{config.DB_TABLE}` "
            f"WHERE (directory, filename) IN ({placeholders})"
        )
        params = [item for pair in chunk for item in pair]
        with conn.cursor() as cur:
            cur.execute(sql, params)
            existing.update(cur.fetchall())
    return existing


HWP_EXTS = {".hwp", ".hwpx"}


def run(csv_path: str, start: int = 0, end=None,
        stop_event: Optional[threading.Event] = None) -> int:
    # 부모가 어떻게 죽든 워커도 같이 죽도록.
    _setup_kill_on_close_job()

    all_rows = []
    with open(csv_path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            all_rows.append(r)

    total_all = len(all_rows)
    # HWP 워커는 .hwp/.hwpx 만 처리 — CSV 에 .pdf 등이 섞여 있어도 무시.
    hwp_rows = [r for r in all_rows if r.get("extension", "").lower() in HWP_EXTS]
    rows = hwp_rows[start:end]
    print(f"  CSV 전체: {total_all:,}건 (그 중 HWP/HWPX {len(hwp_rows):,}건)")
    print(f"  처리 범위: [{start}:{end if end else len(hwp_rows)}] -> {len(rows):,}건")

    if not rows:
        print("  처리할 파일이 없습니다.")
        return 0

    create_db()
    conn = get_conn()
    known_keys = _load_existing_keys(conn, rows)
    if known_keys:
        print(f"  v DB 기존 파일 {len(known_keys):,}건은 파싱 없이 건너뜁니다.")

    err_f = open(ERROR_LOG, "a", newline="", encoding="utf-8-sig")
    err_w = csv.writer(err_f)

    # 이전 실행이 강제 종료돼 남아있을 수 있는 Hwp.exe / HwpFrame.exe
    # 고아 프로세스를 시작 시점에 청소(방어적). 이게 없으면 누적된
    # 좀비들이 메모리·핸들을 잡고 있어 다음 실행이 점점 느려진다.
    _kill_hwp()

    task_q   = mp.Queue()
    result_q = mp.Queue()
    worker   = _spawn_worker(task_q, result_q)
    print(f"  v 워커 프로세스 시작 (PID {worker.pid})")

    pb      = PB(len(rows), offset=start)
    pending = 0

    try:
        for i, row in enumerate(rows):
            if stop_event is not None and stop_event.is_set():
                print("\n\n  중지 요청됨 — 워커 정리 중…")
                break
            d, fn = row["directory"], row["filename"]
            ext   = row.get("extension", "").lower()
            fp    = os.path.join(d, fn)
            key   = (d, fn)

            def _record_error(msg, _d=d, _fn=fn, _ext=ext, _row=row):
                nonlocal pending
                try:
                    with conn.cursor() as cur:
                        cur.execute(INSERT_SQL, (
                            _d, _fn, _ext,
                            _row.get("size_bytes", 0), _row.get("modified", ""),
                            None, "error", msg,
                        ))
                    known_keys.add((_d, _fn))
                    pending += 1
                except Exception:
                    pass

            if key in known_keys:
                pb.tick("skip")
                continue

            if not os.path.exists(fp):
                _record_error("파일 없음")
                pb.tick("error")
                continue

            if len(fp) > 260:
                _record_error(f"경로 초과({len(fp)}자)")
                pb.tick("error")
                continue

            global_idx = start + i
            task_q.put((global_idx, fp, ext))

            text, status, errmsg = None, "success", None
            try:
                _, status, text, errmsg = result_q.get(timeout=config.PARSE_TIMEOUT)
            except Empty:
                status = "error"
                errmsg = "타임아웃/크래시" if worker.is_alive() else "워커 크래시"

                try:
                    worker.kill()
                    worker.join(timeout=5)
                except Exception:
                    pass
                _kill_hwp()
                time.sleep(1)

                task_q   = mp.Queue()
                result_q = mp.Queue()
                worker   = _spawn_worker(task_q, result_q)
                print(f"\n  워커 재시작 (PID {worker.pid}) -- [{global_idx}] {fn}")

                err_w.writerow([d, fn, errmsg])
                _record_error(errmsg)

                if pending >= config.COMMIT_EVERY:
                    conn.commit()
                    pending = 0

                pb.tick("crash")
                continue

            if errmsg:
                err_w.writerow([d, fn, errmsg])

            try:
                with conn.cursor() as cur:
                    cur.execute(INSERT_SQL, (
                        d, fn, ext,
                        row.get("size_bytes", 0), row.get("modified", ""),
                        text, status, errmsg,
                    ))
                known_keys.add(key)
                pending += 1
            except Exception as e:
                err_w.writerow([d, fn, f"DB: {e}"])

            if pending >= config.COMMIT_EVERY:
                conn.commit()
                pending = 0

            pb.tick(status)

    except KeyboardInterrupt:
        print("\n\n  중단됨")
    finally:
        if pending:
            conn.commit()
        try:
            task_q.put(None)
            worker.join(timeout=10)
        except Exception:
            pass
        try:
            worker.kill()
        except Exception:
            pass
        _kill_hwp()

        try:
            with conn.cursor() as cur:
                cur.execute(f"SELECT COUNT(*) FROM `{config.DB_TABLE}`")
                db_count = cur.fetchone()[0]
            match_str = (
                "일치" if pb.cur == len(rows)
                else f"불일치 (처리 {pb.cur} vs 대상 {len(rows)})"
            )
            print(f"\n  [건수 대조]")
            print(f"    이번 배치 대상 : {len(rows):,}건")
            print(f"    이번 배치 처리 : {pb.cur:,}건  {match_str}")
            print(f"    DB 전체 누적   : {db_count:,}건")
        except Exception:
            pass

        conn.close()
        err_f.close()

    pb.done()
    print(f"  에러 로그: {ERROR_LOG}")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Step 2 - HWP 배치 파싱 -> MariaDB",
        epilog="예: python inserter.py --start 0 --end 1000",
    )
    ap.add_argument("--create-db", action="store_true", help="DB/테이블만 생성")
    ap.add_argument("--start",     type=int, default=0,    help="시작 인덱스 (기본: 0)")
    ap.add_argument("--end",       type=int, default=None, help="끝 인덱스 (미지정: 끝까지)")
    ap.add_argument("--csv",       default=config.CSV_FILE,
                    help=f"CSV 경로 (기본: {config.CSV_FILE})")
    args = ap.parse_args()

    print("\n" + "=" * 50)
    print("  Step 2 - HWP Batch Parser -> MariaDB")
    print("=" * 50)
    configure_logging(verbose=False)

    if args.create_db:
        create_db()
        return 0

    if not Path(args.csv).exists():
        print(f"  CSV 없음: {args.csv}")
        return 1

    return run(args.csv, args.start, args.end)


if __name__ == "__main__":
    mp.freeze_support()
    sys.exit(main())
