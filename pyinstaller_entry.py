"""PyInstaller --onefile 빌드용 진입 스크립트.

더블클릭/실행 시 콘솔 메뉴가 아니라 곧장 통합 GUI(탭 4개)가 뜬다.
콘솔 메뉴가 필요하면 `python -m docmine` 또는 소스에서 직접 실행할 것.
"""

import multiprocessing as mp

from docmine.unified_gui import main

if __name__ == "__main__":
    mp.freeze_support()
    main()
