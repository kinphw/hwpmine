"""
PDF 열기/추출 진단 스크립트 (DRM·암호·DRM래퍼 식별용)

사용법:
    같은 폴더에 a.pdf 를 두고
        python test_pdf_open.py
    다른 파일을 보고 싶으면
        python test_pdf_open.py "C:\\path\\to\\some.pdf"

운영환경에서도 그대로 돌려서 로컬 결과와 비교하면 원인이 좁혀집니다:
  - magic 바이트가 %PDF- 가 아니면 → 회사 DRM 솔루션(Fasoo/MarkAny 등)
    이 파일을 wrapping 하고 있을 가능성. 등록되지 않은 프로세스(우리 exe)
    에서는 거부됨.
  - is_encrypted=True / needs_pass=1 → PDF 표준 암호 보호.
  - permissions 가 매우 작은 양수(예: 4, 0) → 추출 차단된 PDF.
  - read_bytes() 실패 → 권한·잠금·OneDrive placeholder 문제.
"""
from __future__ import annotations

import os
import platform
import sys
from pathlib import Path


def diag(path: Path) -> int:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz  # 구버전 호환

    print(f"Python       : {sys.version.split()[0]} ({platform.architecture()[0]})")
    print(f"플랫폼       : {platform.platform()}")
    print(f"PyMuPDF      : {fitz.__version__}")
    print(f"실행 사용자  : {os.getlogin() if hasattr(os, 'getlogin') else '?'}")
    print(f"CWD          : {Path.cwd()}")
    print(f"대상 파일    : {path}")
    print(f"존재 여부    : {path.exists()}")
    if not path.exists():
        return 1
    try:
        size = path.stat().st_size
        print(f"파일 크기    : {size:,} bytes")
    except Exception as e:
        print(f"파일 크기    : ? (stat 실패: {e})")
        size = 0

    # ── 0) magic bytes — DRM wrapper 식별 ────────────────────────
    print()
    print("[0] magic 바이트 (PDF 시그니처 확인)")
    try:
        with open(path, "rb") as f:
            head = f.read(16)
        print(f"  head[0:16] hex   : {head.hex(' ')}")
        print(f"  head[0:8]  ascii : {head[:8]!r}")
        if head.startswith(b"%PDF-"):
            print("  → 정상 PDF 시그니처")
        else:
            print("  ⚠ %PDF- 로 시작하지 않음 — 회사 DRM 으로 wrapping 된")
            print("     컨테이너이거나 손상된 파일일 가능성 높음.")
    except Exception as e:
        print(f"  ✗ 읽기 실패: {type(e).__name__}: {e}")
        print("  → 권한/잠금/OneDrive placeholder 문제 가능.")
        return 2
    print()

    # ── 1) 일반 open ─────────────────────────────────────────────
    print("[1] 기본 open(path) 시도")
    try:
        doc = fitz.open(str(path))
    except Exception as e:
        print(f"  ✗ 실패: {type(e).__name__}: {e}")
        # 바이트로 한 번 더 시도 (외부 잠금/경로 이슈 분리)
        print()
        print("[1b] 바이트로 읽은 뒤 stream 으로 열기")
        try:
            data = path.read_bytes()
            doc = fitz.open(stream=data, filetype="pdf")
            print(f"  ✓ 성공 (페이지 {doc.page_count}쪽)")
            return _try_extract(doc)
        except Exception as e2:
            print(f"  ✗ 실패: {type(e2).__name__}: {e2}")
            return 2

    print(f"  ✓ 열림 (페이지 {doc.page_count}쪽)")
    print()

    # ── 2) 암호/권한 상태 ────────────────────────────────────────
    print("[2] 암호/권한 상태")
    print(f"  needs_pass   : {doc.needs_pass}")
    print(f"  is_encrypted : {doc.is_encrypted}")
    try:
        # permissions 비트마스크 — 페이지 추출/복사/인쇄 허용 여부 표시
        print(f"  permissions  : {doc.permissions} (bitmask)")
    except Exception:
        pass
    try:
        meta = doc.metadata or {}
        print(f"  encryption   : {meta.get('encryption')}")
        print(f"  format       : {meta.get('format')}")
        print(f"  producer     : {meta.get('producer')}")
    except Exception:
        pass
    print()

    # ── 3) 암호 필요 시 빈 패스워드 시도 ─────────────────────────
    if doc.needs_pass:
        print("[3] 빈 패스워드로 authenticate 시도")
        ok = doc.authenticate("")
        print(f"  결과: {ok}  (0=실패, 1=user, 2=owner, -1=취소)")
        if not ok:
            print("  → 추출 권한 없음. DRM/암호 보호된 PDF.")
            return 3

    return _try_extract(doc)


def _try_extract(doc) -> int:
    print()
    print("[4] 텍스트 추출 시도 (첫 페이지)")
    try:
        page = doc[0]
        t_text = page.get_text("text") or ""
        t_blocks = page.get_text("blocks") or []
        print(f"  get_text('text')   : {len(t_text):,}자")
        print(f"  get_text('blocks') : 블록 {len(t_blocks)}개")
        print()
        print("--- 앞 400자 (text 모드) ---")
        print(t_text[:400] if t_text else "(빈 본문 — 스캔본/이미지 PDF 가능성)")
    except Exception as e:
        print(f"  ✗ 추출 실패: {type(e).__name__}: {e}")
        return 4
    finally:
        try:
            doc.close()
        except Exception:
            pass
    return 0


def main() -> int:
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent / "a.pdf"
    rc = diag(target)
    print()
    print(f"종료 코드: {rc}  (0=정상, 1=파일없음, 2=열기실패, 3=DRM/암호, 4=추출실패)")
    return rc


if __name__ == "__main__":
    sys.exit(main())
