"""
Step 1 — HWP/HWPX 파일 스캐너
==============================
드라이브를 재귀 순회하여 .hwp/.hwpx 파일 목록을 CSV로 추출.

단독 실행:
  python scanner.py
  python scanner.py --out my_list.csv
  python scanner.py --drives "C:\\" "D:\\"
"""

import csv
import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

import config

TARGET_EXTENSIONS = {".hwp", ".hwpx"}
SKIP_DIRS = {"$Recycle.Bin", "System Volume Information", "Windows", "ProgramData"}


def scan_files(roots: list[str]) -> list[dict]:
    results = []
    err_count = 0

    for root in roots:
        print(f"\n  [{root}] 스캔 시작...")
        for dirpath, dirnames, filenames in os.walk(root, topdown=True):
            dirnames[:] = [
                d for d in dirnames
                if d not in SKIP_DIRS and not d.startswith("$")
            ]

            for fname in filenames:
                ext = os.path.splitext(fname)[1].lower()
                if ext not in TARGET_EXTENSIONS:
                    continue

                fullpath = os.path.join(dirpath, fname)
                try:
                    stat = os.stat(fullpath)
                    results.append({
                        "directory": dirpath,
                        "filename":  fname,
                        "extension": ext,
                        "size_bytes": stat.st_size,
                        "modified": datetime.fromtimestamp(stat.st_mtime)
                                        .strftime("%Y-%m-%d %H:%M:%S"),
                    })
                except (PermissionError, OSError):
                    err_count += 1

            if results and len(results) % 1000 == 0:
                print(f"    ... {len(results):,}개 발견", flush=True)

    if err_count:
        print(f"\n  ⚠ 접근 불가 파일 {err_count}건 (권한 문제 — 무시됨)")

    return results


def write_csv(rows: list[dict], outpath: Path) -> None:
    fields = ["directory", "filename", "extension", "size_bytes", "modified"]
    with open(outpath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(rows)


def run(drives: list[str] | None = None, out: str | None = None) -> int:
    drives = drives or config.SCAN_DRIVES
    outfile = Path(out or config.CSV_FILE)

    print("=" * 60)
    print("  Step 1 — HWP/HWPX 파일 스캐너")
    print(f"  대상 드라이브: {', '.join(drives)}")
    print(f"  출력 파일    : {outfile.absolute()}")
    print("=" * 60)

    rows = scan_files(drives)

    if not rows:
        print("\n  결과 없음 — HWP/HWPX 파일을 찾지 못했습니다.")
        return 1

    write_csv(rows, outfile)

    hwp_count  = sum(1 for r in rows if r["extension"] == ".hwp")
    hwpx_count = sum(1 for r in rows if r["extension"] == ".hwpx")

    print("\n" + "=" * 60)
    print(f"  완료: 총 {len(rows):,}개 파일")
    print(f"    .hwp  : {hwp_count:,}개")
    print(f"    .hwpx : {hwpx_count:,}개")
    print(f"  저장: {outfile.absolute()}")
    print("=" * 60)
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description="HWP/HWPX 파일 스캐너")
    ap.add_argument("--drives", nargs="+", default=None,
                    help=f"스캔할 드라이브 (기본: {config.SCAN_DRIVES})")
    ap.add_argument("--out", default=None,
                    help=f"출력 CSV 경로 (기본: {config.CSV_FILE})")
    args = ap.parse_args()
    return run(drives=args.drives, out=args.out)


if __name__ == "__main__":
    sys.exit(main())
