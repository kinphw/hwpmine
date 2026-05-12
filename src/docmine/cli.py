"""
Doc Mine — 통합 런처
=====================
스캔·적재는 HWP/PDF 가 별도 CSV·별도 파서로 분리, 검색은 같은 테이블에서
통합 수행. CLI 단축 명령은 HWP 파이프라인만 노출하며, PDF 적재는
`python -m docmine.pdf_inserter` 또는 통합 GUI(`g`) 의 PDF 탭에서.

파이프라인:
  1  스캔   : 드라이브 순회 → CSV 추출   (scanner)        — HWP 전용
  2  적재   : CSV → MariaDB 파싱 적재    (inserter)       — HWP 전용
  3  검색   : DB 기반 GUI 검색기         (search_gui)     — HWP+PDF 통합
  4  추출   : HWP/HWPX → TXT 변환기     (extractor_gui)  — HWP 전용

실행 예:
  docmine            # 대화형 메뉴
  docmine 1          # Step 1만
  docmine all        # 1 → 2 → 3 순차 실행
또는 모듈 호출:
  python -m docmine
"""

import sys
import multiprocessing as mp


BANNER = """\
╔══════════════════════════════════════════╗
║         Doc Mine — 통합 런처             ║
╠══════════════════════════════════════════╣
║  g  통합 GUI (HWP/PDF 스캔·적재 탭 분리) ║
║  1  스캔   HWP 파일 목록 → CSV           ║
║  2  적재   HWP CSV → MariaDB 적재        ║
║  3  검색   GUI 검색기 (HWP+PDF 통합)     ║
║  4  추출   HWP/HWPX → TXT 변환기        ║
║  all       1 → 2 → 3 순차 실행 (HWP)     ║
║  q  종료                                 ║
╚══════════════════════════════════════════╝
  ※ PDF 적재는 통합 GUI(g) 의 PDF 탭 또는
    `python -m docmine.pdf_inserter` 사용."""


def run_step1() -> int:
    from . import scanner, config
    from .drive_picker import pick_drives

    selected = pick_drives(config.SCAN_DRIVES)
    if selected is None:
        print("  스캔 취소됨.")
        return 0
    if not selected:
        print("  선택된 드라이브가 없습니다 — 스캔 취소.")
        return 0
    return scanner.run(drives=selected)


def run_step2() -> int:
    from pathlib import Path
    from . import inserter, config
    if not Path(config.CSV_FILE).exists():
        print(f"  ✗ CSV 없음: {config.CSV_FILE}")
        print("  먼저 Step 1(스캔)을 실행하거나 .env의 CSV_FILE 경로를 확인하세요.")
        return 1
    return inserter.run(config.CSV_FILE, start=0, end=None)


def run_step3() -> int:
    from . import search_gui
    search_gui.main()
    return 0


def run_step4() -> int:
    from . import extractor_gui
    extractor_gui.main()
    return 0


def run_unified_gui() -> int:
    from . import unified_gui
    unified_gui.main()
    return 0


def _step_from_arg(arg: str) -> int | None:
    # 5 = 통합 GUI (g/gui)
    mapping = {
        "1": 1, "2": 2, "3": 3, "4": 4,
        "all": 0, "q": -1,
        "g": 5, "gui": 5,
    }
    return mapping.get(arg.lower())


def main() -> int:
    mp.freeze_support()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass
    if len(sys.argv) > 1:
        choice = sys.argv[1]
    else:
        print(BANNER)
        choice = input("\n  실행할 단계를 입력하세요 (g/1/2/3/4/all/q): ").strip()

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

    if step == 5:
        return run_unified_gui()

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
    sys.exit(main())
