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

try:
    from main import configure_logging
except ImportError:
    raise SystemExit("main.py를 이 스크립트와 같은 폴더에 두세요.")

try:
    import pymysql
except ImportError:
    raise SystemExit("pymysql 필요: pip install pymysql")

import config

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
    parse_status ENUM('success','error','skip') DEFAULT 'success',
    error_msg    VARCHAR(1000),
    parsed_at    DATETIME       DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY   uq_file (directory(500), filename(255))
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci
"""

INSERT_SQL = f"""
INSERT INTO `{config.DB_TABLE}`
    (directory, filename, extension, file_size, file_mtime,
     body_text, parse_status, error_msg)
VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
ON DUPLICATE KEY UPDATE
    body_text=VALUES(body_text), parse_status=VALUES(parse_status),
    error_msg=VALUES(error_msg), parsed_at=CURRENT_TIMESTAMP
"""


def get_conn(use_db: bool = True):
    return pymysql.connect(**config.get_db_config(use_db))


def create_db():
    conn = get_conn(use_db=False)
    with conn.cursor() as cur:
        cur.execute(DDL_DB)
        conn.select_db(config.DB_NAME)
        cur.execute(DDL_TABLE)
    conn.commit()
    conn.close()
    print(f"  v DB [{config.DB_NAME}] / 테이블 [{config.DB_TABLE}] 준비 완료")


# ================================================================
# 진행바
# ================================================================

class PB:
    def __init__(self, total, offset=0):
        self.total = total
        self.offset = offset
        self.cur = self.ok = self.err = self.crash = self.skip = 0
        self.t0 = time.time()

    def tick(self, status="success"):
        self.cur += 1
        if status == "success":
            self.ok += 1
        elif status == "skip":
            self.skip += 1
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
        skip_str = f" skip:{self.skip}" if self.skip else ""
        crash_str = f" crash:{self.crash}" if self.crash else ""
        print(
            f"\r  [{bar}] {idx}/{self.offset + self.total}"
            f"  ok:{self.ok} err:{self.err}{skip_str}{crash_str}"
            f"  ETA {int(eta//60)}:{int(eta%60):02d}  ",
            end="", flush=True,
        )

    def done(self):
        el = time.time() - self.t0
        skip_str = f" skip:{self.skip}" if self.skip else ""
        crash_str = f" crash:{self.crash}" if self.crash else ""
        print(
            f"\n  완료: {int(el//60)}분{int(el%60)}초"
            f"  ok:{self.ok} err:{self.err}{skip_str}{crash_str}"
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


def worker_main(task_q, result_q):
    """
    별도 프로세스로 실행. COM 인스턴스를 유지하며 파일을 파싱.
    크래시 시 이 프로세스만 종료 -> 메인이 새 워커를 띄움.
    """
    from main import ZipDocReader, SectionParser, HWPXDrmError

    zip_reader = ZipDocReader()
    sec_parser = SectionParser()
    com = None
    com_count = 0

    def _get_com():
        nonlocal com
        if com is None:
            import win32com.client as win32
            com = win32.gencache.EnsureDispatch("HwpFrame.HwpObject")
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
        hwp.Open(str(Path(filepath).absolute()))
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


def run(csv_path: str, start: int = 0, end=None) -> int:
    all_rows = []
    with open(csv_path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            all_rows.append(r)

    total_all = len(all_rows)
    rows = all_rows[start:end]
    print(f"  CSV 전체: {total_all:,}건")
    print(f"  처리 범위: [{start}:{end if end else total_all}] -> {len(rows):,}건")

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

    task_q   = mp.Queue()
    result_q = mp.Queue()
    worker   = _spawn_worker(task_q, result_q)
    print(f"  v 워커 프로세스 시작 (PID {worker.pid})")

    pb      = PB(len(rows), offset=start)
    pending = 0

    try:
        for i, row in enumerate(rows):
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
