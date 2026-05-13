"""
PDF → 텍스트 추출 — PyMuPDF(fitz) 기반.

`fitz.Document.get_text("text")` 를 페이지별로 호출해 합친다.
PyMuPDF 는 일반 텍스트 추출 품질·속도가 균형 잡혀 있어 RAG/검색 전처리에 적합.

한계:
- 본문이 이미지로만 들어있는 스캔본 PDF 는 빈 문자열만 나옴 (OCR 미적용).
  → 추후 Tesseract / PaddleOCR 등으로 별도 처리.
- 표 추출은 'text' 모드에서는 단순한 위→아래·좌→우 순서. 표 구조가 필요하면
  fitz 의 'blocks'/'dict' 모드 또는 pdfplumber 로 교체 가능.
"""
from __future__ import annotations

import os
import re
from pathlib import Path


_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _clean(text: str) -> str:
    """제어문자 제거 + 비분리 공백(NBSP) 정리."""
    text = text.replace("\xa0", " ")
    text = _CTRL_RE.sub("", text)
    return text.strip()


_MAX_PATH = 260  # Windows 표준 MAX_PATH (널 포함)


def win_long_path(path: str | Path) -> str:
    """Windows MAX_PATH(260자) 초과 경로만 \\?\\ 접두사로 우회.

    한글·중첩 폴더가 깊은 운영 환경의 긴 경로(>=260자) 처리를 위해
    \\?\\ prefix 를 붙이면 Win32 API 가 ~32,767자까지 허용한다.

    단, prefix 가 붙은 경로는 Win32 의 DOS 경로 정규화 단계를 건너뛰고
    NT object manager 로 직진하기 때문에, 한국 기업 DRM 솔루션
    (Fasoo/MarkAny/Softcamp 등) 의 파일 I/O 후킹 계층이 발동하지 않아
    암호화된 원본 바이트가 그대로 호출자에게 전달될 수 있다. 그래서
    짧은 경로는 손대지 않고, 실제로 260자에 근접·초과하는 경우에만
    prefix 를 적용한다.

    PyMuPDF 의 fitz.open(), os.stat/os.path.exists 양쪽 모두 prefix
    유무에 관계없이 동일하게 받아준다.
    """
    s = str(path)
    if os.name != "nt":
        return s
    p = os.path.abspath(s)
    if p.startswith("\\\\?\\"):
        return p
    # 짧은 경로는 그대로 — DRM 후킹 호환성 우선.
    if len(p) < _MAX_PATH:
        return p
    if p.startswith("\\\\"):  # UNC \\server\share\...
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def _extract_page_blocks(page) -> str:
    """페이지를 'blocks' 모드로 뽑아 y → x 좌표 순으로 정렬해 합친다.

    'text' 모드는 표가 있는 페이지에서 셀이 줄 단위로 섞여 나오는 경우가
    잦은데(예: 헤더 행과 데이터 행이 좌우로 교차), 블록 단위로 받아 좌표
    기준으로 다시 정렬하면 표 영역의 가독성이 눈에 띄게 좋아진다.

    blocks 항목 튜플: (x0, y0, x1, y1, text, block_no, block_type)
    block_type == 0 만 텍스트 블록 (1 은 이미지).
    """
    try:
        blocks = page.get_text("blocks") or []
    except Exception:
        return ""

    text_blocks = [b for b in blocks if len(b) >= 7 and b[6] == 0]
    # 같은 행에 가까운 블록들이 x 순서로 정렬되도록 y0 를 살짝 양자화
    # (소수점 1자리). PDF 의 부동소수 좌표 미세 차이로 행이 어긋나는
    # 현상을 줄임.
    text_blocks.sort(key=lambda b: (round(b[1], 1), round(b[0], 1)))

    return "\n".join(b[4] for b in text_blocks if b[4])


def extract_text(filepath: str | Path) -> str:
    """PDF 파일에서 본문 텍스트를 추출해 정리된 문자열로 반환."""
    import fitz  # PyMuPDF

    fp = win_long_path(filepath)

    parts: list[str] = []
    with fitz.open(fp) as doc:
        for page in doc:
            t = _extract_page_blocks(page)
            if t:
                parts.append(t)
    return _clean("\n".join(parts))
