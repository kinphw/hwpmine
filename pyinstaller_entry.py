"""PyInstaller --onefile 빌드용 진입 스크립트.

`src/hwpmine/__main__.py` 는 패키지 내부에서 `from .cli import main` 를 사용하므로
PyInstaller 가 단독 스크립트로 실행할 때는 relative import 가 깨진다.
이 파일은 패키지 외부에서 절대 import 로 진입하도록 해 준다.

빌드 명령:
    pyinstaller --onefile --name hwpmine --paths src \
        --collect-submodules hwpmine pyinstaller_entry.py
"""

import sys
import multiprocessing as mp

from hwpmine.cli import main

if __name__ == "__main__":
    mp.freeze_support()
    sys.exit(main())
