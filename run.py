"""
개발용 실행 진입점.

실행:
    python run.py            # 대화형 메뉴
    python run.py 3          # Step 3 (검색 GUI) 바로
    python run.py 4          # Step 4 (변환기) 바로
    python run.py all        # 1 → 2 → 3 순차

설치 없이 src-layout 패키지를 바로 실행하기 위해 src/ 를 sys.path 에 추가한다.
배포(PyInstaller) 빌드는 pyinstaller_entry.py 를 사용한다.
"""
import sys
import multiprocessing as mp
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from hwpmine.cli import main  # noqa: E402

if __name__ == "__main__":
    mp.freeze_support()
    sys.exit(main())
