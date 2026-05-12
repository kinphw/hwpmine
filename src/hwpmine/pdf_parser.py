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

import re
from pathlib import Path


_CTRL_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def _clean(text: str) -> str:
    """제어문자 제거 + 비분리 공백(NBSP) 정리."""
    text = text.replace("\xa0", " ")
    text = _CTRL_RE.sub("", text)
    return text.strip()


def extract_text(filepath: str | Path) -> str:
    """PDF 파일에서 본문 텍스트를 추출해 정리된 문자열로 반환."""
    import fitz  # PyMuPDF

    parts: list[str] = []
    with fitz.open(str(filepath)) as doc:
        for page in doc:
            try:
                t = page.get_text("text") or ""
            except Exception:
                t = ""
            if t:
                parts.append(t)
    return _clean("\n".join(parts))
