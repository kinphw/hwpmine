"""
Doc Mine — 통합 GUI 진입점 (.pyw — 콘솔 창 없이 곧장 윈도우만 띄움)

Windows 의 .pyw 확장자는 기본적으로 pythonw.exe 에 연결되어 콘솔 창
없이 실행된다. 더블클릭 또는 `python run.pyw` 양쪽 모두 동일하게
통합 GUI(탭 4개) 가 즉시 뜬다.

콘솔 메뉴 / 단일 단계 실행이 필요하면:
    python -m docmine            # 대화형 메뉴
    python -m docmine 1          # Step 1 (스캔) 만
"""
import sys
import multiprocessing as mp
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from docmine.unified_gui import main  # noqa: E402

if __name__ == "__main__":
    mp.freeze_support()
    main()
