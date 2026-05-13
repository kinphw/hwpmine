"""
PDF 배치 파싱 → MariaDB
========================
CSV 목록(스캔 결과) 에서 .pdf 행만 골라 PyMuPDF 로 본문 텍스트를 뽑아
HWP 와 동일한 `config.DB_TABLE` 에 적재한다.

적재 파이프라인만 HWP 와 분리돼 있고(별도 모듈·별도 CSV·별도 파서),
적재 결과는 단일 테이블로 통합되어 `search_gui` 가 HWP/PDF 를 같이 검색한다.

설계 메모
- 본문 텍스트 추출은 CPU-바운드 — `multiprocessing.Pool` 로 병렬 파싱.
  메인 프로세스는 결과를 받아 DB INSERT 만 수행(직렬, 단일 커넥션).
- 본문이 빈 PDF(스캔본/이미지) 는 `parse_status='empty'` 로 표기 —
  진짜 파싱 실패('error') 와 명시적으로 구분된다.

단독 실행:
  python -m docmine.pdf_inserter
  python -m docmine.pdf_inserter --csv my_pdf_list.csv
  python -m docmine.pdf_inserter --start 0 --end 100
  python -m docmine.pdf_inserter --workers 4
"""
from __future__ import annotations

import argparse
import csv
import multiprocessing as mp
import os
import sys
import threading
from pathlib import Path

from . import config
from .inserter import (
    INSERT_SQL,
    PB,
    _load_existing_keys,
    _setup_kill_on_close_job,
    create_db,
    get_conn,
)
from .pdf_parser import extract_text, win_long_path

PDF_EXTS  = {".pdf"}
ERROR_LOG = "pdf_parse_errors.csv"

# 본문이 비었을 때 사용되는 메시지 — 'error' 가 아니라 'empty' 상태로 기록됨.
EMPTY_TEXT_MSG = "본문 텍스트 없음 (스캔본/이미지 PDF — OCR 미적용)"


# ================================================================
# 워커 — Windows 의 spawn 호환을 위해 모듈 최상위 함수.
# ================================================================

def _extract_pdf_worker(arg):
    """(idx, filepath) → (idx, status, text, errmsg)

    status:
      - "success" : 본문 텍스트 확보
      - "empty"   : 파싱은 성공했지만 본문이 빈 문자열 (스캔본/이미지 PDF)
      - "error"   : 예외 발생 (열기 실패, 손상 등)
    """
    idx, filepath = arg
    try:
        text = extract_text(filepath)
    except Exception as e:
        return idx, "error", None, str(e)[:900]
    if not text:
        return idx, "empty", None, EMPTY_TEXT_MSG
    return idx, "success", text, None


# ================================================================
# 메인
# ================================================================

def run(csv_path: str, start: int = 0, end=None,
        stop_event: threading.Event | None = None) -> int:
    # (A) Job Object — 부모가 어떻게 죽든 워커도 같이 죽도록.
    _setup_kill_on_close_job()

    all_rows = []
    with open(csv_path, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            all_rows.append(r)

    total_all = len(all_rows)
    pdf_rows = [r for r in all_rows if r.get("extension", "").lower() in PDF_EXTS]
    rows = pdf_rows[start:end]
    print(f"  CSV 전체: {total_all:,}건 (그 중 PDF {len(pdf_rows):,}건)")
    print(f"  처리 범위: [{start}:{end if end else len(pdf_rows)}] -> {len(rows):,}건")

    if not rows:
        print("  처리할 PDF 가 없습니다.")
        return 0

    create_db()
    conn = get_conn()
    known_keys = _load_existing_keys(conn, rows)
    if known_keys:
        print(f"  v DB 기존 파일 {len(known_keys):,}건은 파싱 없이 건너뜁니다.")

    err_f = open(ERROR_LOG, "a", newline="", encoding="utf-8-sig")
    err_w = csv.writer(err_f)

    pb = PB(len(rows), offset=start)
    pending = 0

    def _commit_if_due():
        nonlocal pending
        if pending >= config.COMMIT_EVERY:
            conn.commit()
            pending = 0

    def _insert(row, text, status, errmsg):
        """DB 에 1행 INSERT + 에러 로그/카운터 갱신. pending 증가는 여기서."""
        nonlocal pending
        d  = row["directory"]
        fn = row["filename"]
        ext = row.get("extension", "").lower()
        if errmsg and status != "empty":
            # 'empty' 는 정상 케이스로 보고 에러 로그에 남기지 않는다.
            err_w.writerow([d, fn, errmsg])
        try:
            with conn.cursor() as cur:
                cur.execute(INSERT_SQL, (
                    d, fn, ext,
                    row.get("size_bytes", 0), row.get("modified", ""),
                    text, status, errmsg,
                ))
            known_keys.add((d, fn))
            pending += 1
        except Exception as e:
            err_w.writerow([d, fn, f"DB: {e}"])

    # ── 1) 사전 점검(직렬) — DB skip / 파일 없음 ──
    # 경로 260자 초과는 \\?\ prefix 로 우회 → 더 이상 거부하지 않음.
    tasks: list[tuple[int, str]] = []
    for i, row in enumerate(rows):
        if stop_event is not None and stop_event.is_set():
            print("\n  중지 요청됨 (사전 점검 단계).")
            break
        d, fn = row["directory"], row["filename"]
        fp = os.path.join(d, fn)
        if (d, fn) in known_keys:
            pb.tick("skip")
            continue
        if not os.path.exists(win_long_path(fp)):
            _insert(row, None, "error", "파일 없음")
            _commit_if_due()
            pb.tick("error")
            continue
        tasks.append((i, fp))

    # ── 2) 멀티프로세스로 본문 추출, 메인에서 INSERT ──
    if tasks and not (stop_event is not None and stop_event.is_set()):
        workers = max(1, min(config.PDF_WORKERS, len(tasks)))
        cpu = os.cpu_count() or 1
        chunksize = max(1, min(16, len(tasks) // (workers * 8) or 1))
        print(
            f"  PDF 파서 워커 {workers}개로 병렬 파싱"
            f" (감지된 논리 CPU {cpu}개, chunksize={chunksize})…"
        )
        try:
            with mp.Pool(processes=workers) as pool:
                for idx, status, text, errmsg in pool.imap_unordered(
                        _extract_pdf_worker, tasks, chunksize=chunksize):
                    # (B) GUI '중지' 버튼 신호 — with 블록 종료 시
                    # Pool.__exit__ 가 terminate() 를 호출해 워커 즉시 정리.
                    if stop_event is not None and stop_event.is_set():
                        print("\n\n  중지 요청됨 — 워커 종료 중…")
                        break
                    _insert(rows[idx], text, status, errmsg)
                    _commit_if_due()
                    pb.tick(status)
        except KeyboardInterrupt:
            print("\n\n  중단됨")

    # ── 마무리 ──
    try:
        if pending:
            conn.commit()
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
            print(f"    DB 전체 누적   : {db_count:,}건 (HWP + PDF 합산)")
        except Exception:
            pass
    finally:
        conn.close()
        err_f.close()

    pb.done()
    print(f"  에러 로그: {ERROR_LOG}")
    print(f"  (스캔본/이미지 PDF 는 'empty' 로 분류 — 에러 아님)")
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="PDF 배치 파싱 → MariaDB")
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--end",   type=int, default=None)
    ap.add_argument("--csv",   default=config.PDF_CSV_FILE,
                    help=f"CSV 경로 (기본: {config.PDF_CSV_FILE})")
    ap.add_argument("--workers", type=int, default=None,
                    help=f"파서 워커 수 (기본: {config.PDF_WORKERS})")
    args = ap.parse_args()

    if args.workers and args.workers > 0:
        config.PDF_WORKERS = args.workers

    print("\n" + "=" * 50)
    print("  PDF Batch Parser -> MariaDB")
    print("=" * 50)

    if not Path(args.csv).exists():
        print(f"  CSV 없음: {args.csv}")
        return 1

    return run(args.csv, args.start, args.end)


if __name__ == "__main__":
    mp.freeze_support()
    sys.exit(main())
