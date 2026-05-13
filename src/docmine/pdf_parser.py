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


def win_long_path(path: str | Path) -> str:
    """Windows MAX_PATH(260자) 우회용 \\?\\ 접두사 적용.

    한글·중첩 폴더가 깊은 운영 환경에서는 경로가 260자를 쉽게 넘는다.
    Win32 API 는 \\?\\ 접두사가 붙은 절대경로에 대해 ~32,767자까지 허용한다.
    PyMuPDF 의 fitz.open(), os.stat/os.path.exists 양쪽 모두 이 접두사를
    그대로 받아주므로 통일된 진입점에서 변환한다.
    """
    s = str(path)
    if os.name != "nt":
        return s
    p = os.path.abspath(s)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):  # UNC \\server\share\...
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def extract_text(filepath: str | Path) -> str:
    """PDF 파일에서 본문 텍스트를 추출해 정리된 문자열로 반환."""
    import fitz  # PyMuPDF

    fp = win_long_path(filepath)

    parts: list[str] = []
    with fitz.open(fp) as doc:
        for page in doc:
            try:
                t = page.get_text("text") or ""
            except Exception:
                t = ""
            if t:
                parts.append(t)
    return _clean("\n".join(parts))
