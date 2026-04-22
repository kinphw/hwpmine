"""
HWP Mine — 통합 런처
=====================
파이프라인:
  1  스캔   : 드라이브 순회 → CSV 추출   (scanner.py)
  2  적재   : CSV → MariaDB 파싱 적재    (inserter.py)
  3  검색   : DB 기반 GUI 검색기         (search_gui.py)
  4  추출   : HWP/HWPX → TXT 변환기     (extractor_gui.py)

실행 예:
  python run.py          # 대화형 메뉴
  python run.py 1        # Step 1만
  python run.py 2        # Step 2만
  python run.py 3        # Step 3만
  python run.py 4        # Step 4만
  python run.py all      # 1 → 2 → 3 순차 실행
"""

import sys
import multiprocessing as mp


BANNER = """\
╔══════════════════════════════════════════╗
║         HWP Mine — 통합 런처             ║
╠══════════════════════════════════════════╣
║  1  스캔   HWP 파일 목록 → CSV           ║
║  2  적재   CSV → MariaDB 파싱 적재       ║
║  3  검색   GUI 검색기 실행               ║
║  4  추출   HWP/HWPX → TXT 변환기        ║
║  all       1 → 2 → 3 순차 실행           ║
║  q  종료                                 ║
╚══════════════════════════════════════════╝"""


def run_step1() -> int:
    import scanner
    return scanner.run()


def run_step2() -> int:
    import inserter
    from pathlib import Path
    import config
    if not Path(config.CSV_FILE).exists():
        print(f"  ✗ CSV 없음: {config.CSV_FILE}")
        print("  먼저 Step 1(스캔)을 실행하거나 .env의 CSV_FILE 경로를 확인하세요.")
        return 1
    return inserter.run(config.CSV_FILE, start=0, end=None)


def run_step3() -> int:
    import search_gui
    search_gui.main()
    return 0


def run_step4() -> int:
    import extractor_gui
    extractor_gui.main()
    return 0


def _step_from_arg(arg: str) -> int | None:
    mapping = {"1": 1, "2": 2, "3": 3, "4": 4, "all": 0, "q": -1}
    return mapping.get(arg.lower())


def main() -> int:
    if len(sys.argv) > 1:
        choice = sys.argv[1]
    else:
        print(BANNER)
        choice = input("\n  실행할 단계를 입력하세요 (1/2/3/all/q): ").strip()

    step = _step_from_arg(choice)

    if step == -1 or step is None:
        if choice.lower() in ("q", "quit", "exit"):
            return 0
        print(f"  알 수 없는 입력: {choice!r}")
        return 1

    if step == 1:
        return run_step1()

    if step == 2:
        return run_step2()

    if step == 3:
        return run_step3()

    if step == 4:
        return run_step4()

    if step == 0:   # all
        rc = run_step1()
        if rc != 0:
            print("  Step 1 실패 — 중단합니다.")
            return rc
        rc = run_step2()
        if rc != 0:
            print("  Step 2 실패 — 중단합니다.")
            return rc
        return run_step3()

    return 0


if __name__ == "__main__":
    mp.freeze_support()   # Windows exe 패키징 대비
    sys.exit(main())
