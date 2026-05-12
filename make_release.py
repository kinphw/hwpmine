"""배포용 zip 생성 스크립트 (단일 실행파일 버전).

사용법:
    pyinstaller --onefile --name docmine --paths src \
        --collect-submodules docmine pyinstaller_entry.py
    python make_release.py

수행 작업:
    1) src/docmine/__init__.py 의 버전 읽기
    2) dist/docmine.exe 가 있는지 확인
    3) release/docmine_v<버전>/ 에 exe, install.bat, .env.example, README.md 모음
    4) docmine_v<버전>.zip 으로 압축
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
INIT_PY = ROOT / "src" / "docmine" / "__init__.py"
DIST = ROOT / "dist"
RELEASE = ROOT / "release"
EXE_NAME = "docmine.exe"

INCLUDE = ["install.bat", ".env.example", "README.md"]


def read_version() -> str:
    text = INIT_PY.read_text(encoding="utf-8")
    m = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', text, re.M)
    if not m:
        sys.exit(f"[오류] 버전을 찾을 수 없습니다: {INIT_PY}")
    return m.group(1)


def find_exe() -> Path:
    exe = DIST / EXE_NAME
    if not exe.exists():
        sys.exit(
            f"[오류] 실행파일이 없습니다: {exe}\n"
            f"       먼저 다음 명령으로 빌드하세요:\n"
            f"         pyinstaller --onefile --name docmine --paths src "
            f"--collect-submodules docmine pyinstaller_entry.py"
        )
    return exe


def stage(version: str, exe: Path) -> Path:
    stage_dir = RELEASE / f"docmine_v{version}"
    if stage_dir.exists():
        shutil.rmtree(stage_dir)
    stage_dir.mkdir(parents=True)

    shutil.copy2(exe, stage_dir / exe.name)
    for name in INCLUDE:
        src = ROOT / name
        if not src.exists():
            sys.exit(f"[오류] 필수 파일이 없습니다: {src}")
        shutil.copy2(src, stage_dir / name)
    return stage_dir


def make_zip(stage_dir: Path, version: str) -> Path:
    zip_path = ROOT / f"docmine_v{version}.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in stage_dir.rglob("*"):
            if p.is_file():
                zf.write(p, arcname=p.relative_to(stage_dir))
    return zip_path


def main() -> None:
    version = read_version()
    print(f"버전: {version}")

    exe = find_exe()
    stage_dir = stage(version, exe)
    zip_path = make_zip(stage_dir, version)

    print()
    print("=" * 60)
    print(f"  완료: {zip_path.name}")
    print(f"  경로: {zip_path}")
    print(f"  내용: {exe.name}, " + ", ".join(INCLUDE))
    print("=" * 60)


if __name__ == "__main__":
    main()
