"""배포용 zip 생성 스크립트.

사용법:
    python -m build --wheel     # 먼저 wheel 을 빌드해두고
    python make_release.py      # 그 다음 실행

수행 작업:
    1) src/hwpmine/__init__.py 의 버전 읽기
    2) dist/ 에서 해당 버전의 wheel 찾기
    3) release/hwpmine_v<버전>/ 에 whl, install.bat, .env.example, README.md 모음
    4) hwpmine_v<버전>.zip 으로 압축
"""
from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
INIT_PY = ROOT / "src" / "hwpmine" / "__init__.py"
DIST = ROOT / "dist"
RELEASE = ROOT / "release"

INCLUDE = ["install.bat", ".env.example", "README.md"]


def read_version() -> str:
    text = INIT_PY.read_text(encoding="utf-8")
    m = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', text, re.M)
    if not m:
        sys.exit(f"[오류] 버전을 찾을 수 없습니다: {INIT_PY}")
    return m.group(1)


def find_wheel(version: str) -> Path:
    wheel = DIST / f"hwpmine-{version}-py3-none-any.whl"
    if not wheel.exists():
        sys.exit(
            f"[오류] wheel 이 없습니다: {wheel}\n"
            f"       먼저 'python -m build --wheel' 로 빌드하세요."
        )
    return wheel


def stage(version: str, wheel: Path) -> Path:
    stage_dir = RELEASE / f"hwpmine_v{version}"
    if stage_dir.exists():
        shutil.rmtree(stage_dir)
    stage_dir.mkdir(parents=True)

    shutil.copy2(wheel, stage_dir / wheel.name)
    for name in INCLUDE:
        src = ROOT / name
        if not src.exists():
            sys.exit(f"[오류] 필수 파일이 없습니다: {src}")
        shutil.copy2(src, stage_dir / name)
    return stage_dir


def make_zip(stage_dir: Path, version: str) -> Path:
    zip_path = ROOT / f"hwpmine_v{version}.zip"
    if zip_path.exists():
        zip_path.unlink()
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in stage_dir.rglob("*"):
            if p.is_file():
                zf.write(p, arcname=p.relative_to(stage_dir.parent))
    return zip_path


def main() -> None:
    version = read_version()
    print(f"버전: {version}")

    wheel = find_wheel(version)
    stage_dir = stage(version, wheel)
    zip_path = make_zip(stage_dir, version)

    print()
    print("=" * 60)
    print(f"  완료: {zip_path.name}")
    print(f"  경로: {zip_path}")
    print(f"  내용: {wheel.name}, " + ", ".join(INCLUDE))
    print("=" * 60)


if __name__ == "__main__":
    main()
