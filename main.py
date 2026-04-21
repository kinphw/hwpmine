"""
HWPParser - HWPX / HWP 파일 파싱 라이브러리
============================================
HWPX(한글 XML 형식) 및 DRM 보호 파일에서 텍스트/구조를 추출하는 OOP 기반 모듈.

백엔드 선택 전략
----------------
  backend="auto"  (기본값)
      1차: 직접 ZIP 파싱 (의존성 없음, 빠름)
      2차: ZIP 실패 시(DRM 등) → COM 백엔드로 자동 전환
  backend="zip"   ZIP 직접 파싱만 사용 (DRM 파일 불가)
  backend="com"   win32com HWP 자동화만 사용 (Windows + HWP 설치 필수)

COM 백엔드 동작 원리 (DRM 우회)
---------------------------------
  DRM이 걸린 HWPX는 ZIP 헤더 자체가 암호화되어 있어 직접 열기 불가.
  COM 백엔드는 실제 HWP 프로그램을 비가시(invisible) 모드로 띄운 뒤
  HWP가 DRM을 자체 해제하도록 하고, InitScan/GetText API로 텍스트를
  스트리밍 방식으로 수집합니다. (hwp_auto.py 방식 참조)

  GetText() 상태 코드:
      1   일반 텍스트
      2   컨트롤(표·그림) 진입
      3   컨트롤 탈출
      4   필드 시작
      5   필드 끝
      101 문서 끝

HWPX ZIP 파일 구조:
  - Contents/section0.xml ...  본문 섹션
  - Contents/header.xml        문서 헤더/스타일
  - Contents/content.hpf       패키지 목록

XML 주요 네임스페이스:
  - hp: http://www.hancom.co.kr/hwpml/2012/paragraph
  - hc: http://www.hancom.co.kr/hwpml/2012/core
"""

from __future__ import annotations

import logging
import re
import zipfile
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum, auto
from pathlib import Path
from typing import Iterable, Iterator, List, Optional
from xml.etree import ElementTree as ET

LOGGER = logging.getLogger("HWPParser")


# ---------------------------------------------------------------------------
# 예외 클래스
# ---------------------------------------------------------------------------

class HWPXError(Exception):
    """HWPParser 최상위 예외"""


class HWPXFormatError(HWPXError):
    """유효하지 않은 HWPX 파일 형식"""


class HWPXParseError(HWPXError):
    """XML 파싱 중 오류"""


class HWPXDrmError(HWPXError):
    """DRM 보호 파일 — ZIP 직접 파싱 불가, COM 백엔드 필요"""


class HWPXComError(HWPXError):
    """COM 백엔드 초기화/실행 오류 (win32com 미설치 또는 HWP 미설치)"""


# ---------------------------------------------------------------------------
# 백엔드 선택 Enum
# ---------------------------------------------------------------------------

class Backend(Enum):
    AUTO = auto()   # ZIP 먼저, 실패 시 COM
    ZIP  = auto()   # ZIP 직접 파싱만
    COM  = auto()   # COM(win32com) 자동화만


# ---------------------------------------------------------------------------
# 상수
# ---------------------------------------------------------------------------

# HWPX XML 네임스페이스 (로컬 이름 기반 매칭에도 사용)
HWPX_NS = {
    "hp": "http://www.hancom.co.kr/hwpml/2012/paragraph",
    "hc": "http://www.hancom.co.kr/hwpml/2012/core",
    "hh": "http://www.hancom.co.kr/hwpml/2012/head",
    "hf": "http://www.hancom.co.kr/hwpml/2012/history",
    "hs": "http://www.hancom.co.kr/hwpml/2012/section",
}

# section 파일 패턴
SECTION_FILE_PATTERN = re.compile(
    r"^(Contents/)?[Ss]ection(\d+)\.xml$"
)

CONTENT_HPF_PATHS = [
    "Contents/content.hpf",
    "content.hpf",
]


# ---------------------------------------------------------------------------
# 텍스트 추출 옵션 (GPT 버전 참조)
# ---------------------------------------------------------------------------

@dataclass
class TextExtractionOptions:
    """
    텍스트 추출 동작을 제어하는 옵션 객체.

    Attributes:
        preserve_blank_lines:  빈 줄 유지 여부 (False면 빈 줄 제거)
        normalize_whitespace:  내부 공백 정규화 (탭·반복 공백 → 단일 공백)
        strip_lines:           각 줄 앞뒤 공백 제거
        include_tables:        표 내용 포함 여부
        section_separator:     섹션 간 구분자
    """
    preserve_blank_lines: bool = True
    normalize_whitespace: bool = True
    strip_lines: bool = True
    include_tables: bool = True
    section_separator: str = "\n\n"


# ---------------------------------------------------------------------------
# 데이터 모델 (불변 값 객체)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class TextRun:
    """단일 텍스트 런 (서식 단위)"""
    text: str

    def __str__(self) -> str:
        return self.text


@dataclass
class Paragraph:
    """단락"""
    runs: list[TextRun] = field(default_factory=list)
    style_id: Optional[str] = None     # 스타일 참조 ID
    level: int = 0                      # 개요 수준

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    def is_empty(self) -> bool:
        return not self.text.strip()

    def __str__(self) -> str:
        return self.text


@dataclass
class TableCell:
    """표 셀"""
    paragraphs: list[Paragraph] = field(default_factory=list)
    row_span: int = 1
    col_span: int = 1

    @property
    def text(self) -> str:
        return "\n".join(p.text for p in self.paragraphs)


@dataclass
class TableRow:
    """표 행"""
    cells: list[TableCell] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\t".join(c.text for c in self.cells)


@dataclass
class Table:
    """표"""
    rows: list[TableRow] = field(default_factory=list)

    @property
    def text(self) -> str:
        return "\n".join(r.text for r in self.rows)

    def to_plain_text(self, cell_sep: str = " | ", row_sep: str = "\n") -> str:
        parts = []
        for row in self.rows:
            parts.append(cell_sep.join(c.text.replace("\n", " ") for c in row.cells))
        return row_sep.join(parts)


# 섹션 내 블록 타입
Block = Paragraph | Table


@dataclass
class Section:
    """문서 섹션 (본문, 머리말, 꼬리말 등)"""
    index: int
    blocks: list[Block] = field(default_factory=list)

    @property
    def paragraphs(self) -> list[Paragraph]:
        return [b for b in self.blocks if isinstance(b, Paragraph)]

    @property
    def tables(self) -> list[Table]:
        return [b for b in self.blocks if isinstance(b, Table)]

    @property
    def text(self) -> str:
        lines: list[str] = []
        for block in self.blocks:
            if isinstance(block, Paragraph):
                lines.append(block.text)
            elif isinstance(block, Table):
                lines.append(block.to_plain_text())
        return "\n".join(lines)


@dataclass
class HWPXDocument:
    """파싱된 HWPX 문서 전체"""
    path: Path
    sections: list[Section] = field(default_factory=list)
    metadata: dict[str, str] = field(default_factory=dict)

    # ------------------------------------------------------------------
    # 텍스트 추출 (공개 API)
    # ------------------------------------------------------------------

    @property
    def text(self) -> str:
        """섹션 전체를 합친 순수 텍스트"""
        return "\n\n".join(s.text for s in self.sections)

    def iter_paragraphs(self) -> Iterator[Paragraph]:
        """모든 섹션의 단락을 순서대로 순회"""
        for section in self.sections:
            yield from section.paragraphs

    def iter_tables(self) -> Iterator[Table]:
        """모든 섹션의 표를 순서대로 순회"""
        for section in self.sections:
            yield from section.tables

    def extract_text(
        self,
        options: Optional[TextExtractionOptions] = None,
        *,
        # 하위 호환 개별 kwargs (options 우선)
        include_tables: Optional[bool] = None,
        skip_empty: Optional[bool] = None,
        section_separator: Optional[str] = None,
        paragraph_separator: str = "\n",
    ) -> str:
        """
        텍스트 추출.

        Args:
            options:             TextExtractionOptions 인스턴스 (None이면 기본값)
            include_tables:      표 내용 포함 여부 (options 없을 때 개별 지정)
            skip_empty:          빈 단락 제외 여부 (options 없을 때 개별 지정)
            section_separator:   섹션 구분자 (options 없을 때 개별 지정)
            paragraph_separator: 단락 구분자
        """
        opt = options or TextExtractionOptions()

        # 개별 kwargs가 명시된 경우 options보다 우선
        _include_tables  = include_tables  if include_tables  is not None else opt.include_tables
        _preserve_blanks = not skip_empty  if skip_empty      is not None else opt.preserve_blank_lines
        _section_sep     = section_separator if section_separator is not None else opt.section_separator

        section_texts: list[str] = []
        for section in self.sections:
            lines: list[str] = []
            for block in section.blocks:
                if isinstance(block, Paragraph):
                    lines.append(block.text)
                elif isinstance(block, Table) and _include_tables:
                    lines.append(block.to_plain_text())
            raw = paragraph_separator.join(lines)
            processed = _postprocess_lines(
                raw.splitlines(),
                preserve_blank_lines=_preserve_blanks,
                normalize_whitespace=opt.normalize_whitespace,
                strip_lines=opt.strip_lines,
            )
            section_texts.append(processed)

        return _section_sep.join(t for t in section_texts if t)

    def __repr__(self) -> str:
        return (
            f"HWPXDocument(path={self.path.name!r}, "
            f"sections={len(self.sections)}, "
            f"paragraphs={sum(len(s.paragraphs) for s in self.sections)})"
        )


# ---------------------------------------------------------------------------
# XML 헬퍼
# ---------------------------------------------------------------------------

def _local(tag: str) -> str:
    """'{namespace}localname' → 'localname'"""
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag


def _find_all_by_local(element: ET.Element, *local_names: str) -> Iterator[ET.Element]:
    """네임스페이스 무관하게 로컬 이름으로 하위 요소 재귀 탐색"""
    target = set(local_names)
    for child in element:
        if _local(child.tag) in target:
            yield child
        yield from _find_all_by_local(child, *local_names)


def _iter_direct_by_local(element: ET.Element, *local_names: str) -> Iterator[ET.Element]:
    """직계 자식 중 로컬 이름 매칭"""
    target = set(local_names)
    for child in element:
        if _local(child.tag) in target:
            yield child


def _postprocess_lines(
    lines: Iterable[str],
    *,
    preserve_blank_lines: bool = True,
    normalize_whitespace: bool = True,
    strip_lines: bool = True,
) -> str:
    """
    줄 목록을 후처리하여 최종 텍스트 반환. (GPT 버전의 _postprocess_lines 참조)

    - normalize_whitespace: 탭·반복 공백 → 단일 공백
    - strip_lines:          줄 앞뒤 공백 제거
    - preserve_blank_lines: 빈 줄 유지 여부 (False면 제거)
    - 연속 3개 이상 빈 줄 → 최대 2개로 축소
    """
    processed: List[str] = []
    for line in lines:
        if normalize_whitespace:
            line = re.sub(r"[ \t\r\f\v]+", " ", line)
        if strip_lines:
            line = line.strip()
        if line:
            processed.append(line)
        elif preserve_blank_lines:
            processed.append("")

    text = "\n".join(processed)

    if preserve_blank_lines:
        text = re.sub(r"\n{3,}", "\n\n", text)   # 연속 빈 줄 최대 2개
    else:
        text = re.sub(r"\n{2,}", "\n", text)      # 빈 줄 전부 제거

    return text.strip()


# ---------------------------------------------------------------------------
# 노드 파서 (확장 지점)
# ---------------------------------------------------------------------------

class BaseNodeParser(ABC):
    """XML 요소 파서 추상 기반 클래스.

    새로운 블록 유형(주석, 그림 등)을 추가하려면 이 클래스를 상속하고
    SectionParser.register()로 등록하면 됩니다.
    """

    @abstractmethod
    def can_parse(self, local_tag: str) -> bool:
        """이 파서가 처리할 수 있는 태그 이름인지 확인"""

    @abstractmethod
    def parse(self, element: ET.Element) -> Optional[Block]:
        """요소를 Block(Paragraph 또는 Table)으로 변환. 무시할 경우 None 반환"""


class ParagraphParser(BaseNodeParser):
    """<hp:para> / <para> / <p> 파싱"""

    _PARA_TAGS = {"para", "p", "PARA", "P"}
    # tail 텍스트를 수집할 의미 있는 태그 (공백 전용 tail 제외)
    _TEXT_TAGS = {"t", "T", "text"}

    def can_parse(self, local_tag: str) -> bool:
        return local_tag in self._PARA_TAGS

    def parse(self, element: ET.Element) -> Optional[Paragraph]:
        runs: list[TextRun] = []

        # 스타일 속성
        style_id = element.get("styleIDRef") or element.get("styleId")
        level_str = element.get("outlineLevel", "0")
        try:
            level = int(level_str)
        except (ValueError, TypeError):
            level = 0

        # ── 1차: <run> 하위의 <t> 요소 수집 ──────────────────────────
        for run_el in _find_all_by_local(element, "run", "Run", "RUN"):
            text_parts: list[str] = []
            for t_el in _iter_direct_by_local(run_el, "t", "T"):
                text_parts.append(t_el.text or "")
                # tail 텍스트 수집 (GPT 버전 참조: 일부 생성기가 tail에 텍스트를 씀)
                if t_el.tail and t_el.tail.strip():
                    text_parts.append(t_el.tail)
            # 줄 바꿈 요소 처리
            for _ in _iter_direct_by_local(run_el, "lineBreak", "LineBreak"):
                text_parts.append("\n")
            if text_parts:
                runs.append(TextRun("".join(text_parts)))

        # ── 2차 fallback: <run> 없이 바로 존재하는 <t> 수집 ──────────
        if not runs:
            direct_text = "".join(
                (t.text or "") for t in _iter_direct_by_local(element, "t", "T")
            )
            if direct_text:
                runs.append(TextRun(direct_text))

        # ── 3차 fallback: 구조가 전혀 다른 변형 포맷 ─────────────────
        # <t>/<text> 조차 없을 때 전체 iter()로 텍스트 긁기 (GPT 버전 참조)
        if not runs:
            chunks: list[str] = []
            for elem in element.iter():
                local = _local(elem.tag)
                if local in self._TEXT_TAGS and elem.text:
                    chunks.append(elem.text)
                if elem.tail and elem.tail.strip() and elem is not element:
                    chunks.append(elem.tail)
            if chunks:
                runs.append(TextRun("".join(chunks)))

        return Paragraph(runs=runs, style_id=style_id, level=level)


class TableParser(BaseNodeParser):
    """<hp:tbl> / <tbl> / <table> 파싱"""

    _TBL_TAGS = {"tbl", "table", "Tbl", "Table", "TBL", "TABLE"}

    def __init__(self) -> None:
        self._para_parser = ParagraphParser()

    def can_parse(self, local_tag: str) -> bool:
        return local_tag in self._TBL_TAGS

    def parse(self, element: ET.Element) -> Optional[Table]:
        rows: list[TableRow] = []

        for tr_el in _find_all_by_local(element, "tr", "Tr", "TR"):
            cells: list[TableCell] = []
            for tc_el in _iter_direct_by_local(tr_el, "tc", "Tc", "TC"):
                row_span = int(tc_el.get("rowSpan", "1") or 1)
                col_span = int(tc_el.get("colSpan", "1") or 1)
                paragraphs: list[Paragraph] = []
                for para_el in _find_all_by_local(tc_el, "para", "p", "PARA", "P"):
                    para = self._para_parser.parse(para_el)
                    if para is not None:
                        paragraphs.append(para)
                cells.append(TableCell(
                    paragraphs=paragraphs,
                    row_span=row_span,
                    col_span=col_span,
                ))
            if cells:
                rows.append(TableRow(cells=cells))

        return Table(rows=rows) if rows else None


# ---------------------------------------------------------------------------
# 섹션 파서
# ---------------------------------------------------------------------------

class SectionParser:
    """section XML 요소를 Section 객체로 변환.

    파서 레지스트리를 통해 새로운 블록 파서를 등록할 수 있습니다.
    """

    def __init__(self) -> None:
        self._parsers: list[BaseNodeParser] = [
            ParagraphParser(),
            TableParser(),
        ]

    def register(self, parser: BaseNodeParser) -> None:
        """새 노드 파서 등록 (우선순위 낮음: 뒤에 추가됨)"""
        self._parsers.append(parser)

    def register_first(self, parser: BaseNodeParser) -> None:
        """새 노드 파서 등록 (우선순위 높음: 앞에 추가됨)"""
        self._parsers.insert(0, parser)

    def parse_element(self, element: ET.Element, index: int) -> Section:
        """ET.Element → Section"""
        section = Section(index=index)
        self._collect_blocks(element, section.blocks)
        return section

    def parse_xml(self, xml_bytes: bytes, index: int) -> Section:
        """XML 바이트 → Section"""
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError as exc:
            raise HWPXParseError(f"섹션 {index} XML 파싱 실패: {exc}") from exc
        return self.parse_element(root, index)

    # ------------------------------------------------------------------
    # 내부 구현
    # ------------------------------------------------------------------

    def _collect_blocks(self, element: ET.Element, blocks: list[Block]) -> None:
        """요소 트리를 순회하며 블록 수집 (재귀)"""
        for child in element:
            local = _local(child.tag)
            block = self._dispatch(local, child)
            if block is not None:
                blocks.append(block)
            else:
                # 알 수 없는 컨테이너 요소는 재귀 탐색
                if local not in {"para", "p", "tbl", "table", "PARA", "P", "TBL", "TABLE"}:
                    self._collect_blocks(child, blocks)

    def _dispatch(self, local_tag: str, element: ET.Element) -> Optional[Block]:
        for parser in self._parsers:
            if parser.can_parse(local_tag):
                return parser.parse(element)
        return None


# ---------------------------------------------------------------------------
# 문서 리더 추상 기반 (백엔드 전략 패턴)
# ---------------------------------------------------------------------------

class BaseDocReader(ABC):
    """문서 읽기 백엔드 추상 클래스."""

    @abstractmethod
    def read_document(self, path: Path, section_parser: "SectionParser") -> HWPXDocument:
        """파일을 읽어 HWPXDocument 반환"""


# ---------------------------------------------------------------------------
# ZIP 백엔드 — 직접 파싱 (비DRM)
# ---------------------------------------------------------------------------

class ZipDocReader(BaseDocReader):
    """HWPX(ZIP) 아카이브 직접 파싱. DRM 없는 파일 전용."""

    # --- 내부 헬퍼 ---

    @staticmethod
    def _is_drm_protected(path: Path) -> bool:
        """파일이 DRM 암호화된 ZIP인지 빠르게 판별.

        HWPX DRM 파일은 ZIP 시그니처(PK\\x03\\x04) 대신
        한컴 자체 암호화 헤더로 시작합니다.
        """
        try:
            with open(path, "rb") as f:
                header = f.read(4)
            # 정상 ZIP: 50 4B 03 04
            return header[:2] != b"PK"
        except OSError:
            return False

    def read_document(self, path: Path, section_parser: "SectionParser") -> HWPXDocument:
        # 확장자 검증 (GPT 버전 참조)
        if path.suffix.lower() not in {".hwpx", ".hwp"}:
            raise HWPXFormatError(
                f"{path.name}: 지원하지 않는 파일 형식입니다 "
                f"({path.suffix}). .hwpx 또는 .hwp 파일을 사용하세요."
            )

        # DRM 여부를 먼저 파일 헤더로 확인
        if self._is_drm_protected(path):
            raise HWPXDrmError(
                f"{path.name}: DRM 보호 파일입니다. "
                "backend='com' 또는 backend='auto'를 사용하세요."
            )

        try:
            zf = zipfile.ZipFile(path, "r")
        except zipfile.BadZipFile as exc:
            raise HWPXDrmError(
                f"{path.name}: ZIP 열기 실패 — DRM 보호 가능성이 있습니다."
            ) from exc

        with zf:
            entries = self._section_entries(zf)
            if not entries:
                raise HWPXFormatError(
                    f"{path.name}: 섹션 파일이 없습니다. HWPX 형식을 확인하세요."
                )
            metadata = self._read_metadata(zf)
            sections: list[Section] = []
            for idx, entry_name in entries:
                LOGGER.debug("섹션 읽는 중: %s", entry_name)
                xml_bytes = zf.read(entry_name)
                section = section_parser.parse_xml(xml_bytes, idx)
                sections.append(section)

        LOGGER.debug(
            "%s: %d개 섹션, %d개 단락 파싱 완료",
            path.name,
            len(sections),
            sum(len(s.paragraphs) for s in sections),
        )
        return HWPXDocument(path=path, sections=sections, metadata=metadata)

    @staticmethod
    def _section_entries(zf: zipfile.ZipFile) -> list[tuple[int, str]]:
        result: list[tuple[int, str]] = []
        for name in zf.namelist():
            m = SECTION_FILE_PATTERN.match(name)
            if m:
                result.append((int(m.group(2)), name))
        result.sort(key=lambda x: x[0])
        return result

    @staticmethod
    def _read_metadata(zf: zipfile.ZipFile) -> dict[str, str]:
        meta: dict[str, str] = {}
        for hpf_path in CONTENT_HPF_PATHS:
            if hpf_path in zf.namelist():
                try:
                    data = zf.read(hpf_path)
                    root = ET.fromstring(data)
                    for child in root.iter():
                        local = _local(child.tag)
                        if local in ("title", "creator", "description", "subject", "language"):
                            if child.text:
                                meta[local] = child.text.strip()
                except (ET.ParseError, KeyError):
                    pass
                break
        return meta


# ---------------------------------------------------------------------------
# COM 백엔드 — win32com HWP 자동화 (DRM 우회)
# ---------------------------------------------------------------------------

class ComDocReader(BaseDocReader):
    """
    win32com을 통해 실제 HWP 프로세스를 제어하여 텍스트 추출.
    DRM 보호 파일에서도 동작합니다. (Windows + 한글 설치 필수)

    hwp_auto.py 참조:
        hwp = win32.gencache.EnsureDispatch("HwpFrame.HwpObject")
        hwp.Open(path)
        hwp.InitScan(Range=0xff)
        state, text = hwp.GetText()
        hwp.ReleaseScan()

    GetText() 상태 코드:
        1   일반 텍스트 (단락 끝은 \\r 포함)
        2   컨트롤(표·그림 등) 진입
        3   컨트롤 탈출
        4   필드 시작
        5   필드 끝
        101 문서 끝
    """

    # 컨트롤 진입/탈출 상태 코드
    _STATE_TEXT         = 1
    _STATE_CTRL_IN      = 2
    _STATE_CTRL_OUT     = 3
    _STATE_FIELD_START  = 4
    _STATE_FIELD_END    = 5
    _STATE_END          = 101

    def __init__(self, visible: bool = False) -> None:
        """
        Args:
            visible: HWP 창을 화면에 표시할지 여부.
                     False(기본값)이면 백그라운드에서 실행.
                     일부 DRM 환경에서는 True로 설정해야 할 수 있음.
        """
        self.visible = visible

    def read_document(self, path: Path, section_parser: "SectionParser") -> HWPXDocument:
        import threading, time

        stop_event = threading.Event()

        def _loop():
            try:
                import win32gui, win32con  # type: ignore
            except ImportError:
                return
            BUTTONS = ["접근 허용(&A)", "접근 허용", "확인(&O)", "확인", "OK"]
            while not stop_event.is_set():
                try:
                    def _on_win(hwnd, _):
                        if not win32gui.IsWindowVisible(hwnd):
                            return
                        def _on_child(child, _):
                            try:
                                if win32gui.GetClassName(child) != "Button":
                                    return
                                txt = win32gui.GetWindowText(child)
                                if any(b == txt or b in txt for b in BUTTONS):
                                    win32gui.SendMessage(child, win32con.BM_CLICK, 0, 0)
                            except Exception:
                                pass
                        try:
                            win32gui.EnumChildWindows(hwnd, _on_child, None)
                        except Exception:
                            pass
                    win32gui.EnumWindows(_on_win, None)
                except Exception:
                    pass
                time.sleep(0.3)

        t = threading.Thread(target=_loop, daemon=True)
        t.start()

        try:
            hwp = self._open_hwp(path)
            try:
                sections = self._extract_sections(hwp)
            finally:
                self._close_hwp(hwp)
                time.sleep(2)   # close 후 팝업 대기
        finally:
            stop_event.set()

        return HWPXDocument(path=path, sections=sections, metadata={})

    @staticmethod
    def _start_popup_dismisser():
        pass  # read_document 내부에서 직접 처리

    # ------------------------------------------------------------------
    # HWP COM 생명주기
    # ------------------------------------------------------------------

    def _open_hwp(self, path: Path):
        """HWP COM 객체 생성 및 파일 열기."""
        try:
            import win32com.client as win32  # type: ignore
        except ImportError as exc:
            raise HWPXComError(
                "win32com을 찾을 수 없습니다. 'pip install pywin32'로 설치하세요."
            ) from exc

        try:
            hwp = win32.gencache.EnsureDispatch("HwpFrame.HwpObject")
        except Exception as exc:
            raise HWPXComError(
                "HWP COM 객체 생성 실패. 한글(HWP)이 설치되어 있는지 확인하세요."
            ) from exc

        abs_path = str(path.absolute())

        try:
            hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        except Exception:
            pass

        try:
            hwp.Open(abs_path)
        except Exception as exc:
            raise HWPXComError(f"파일 열기 실패: {path.name} — {exc}") from exc

        hwp.XHwpWindows.Item(0).Visible = self.visible
        return hwp

    @staticmethod
    def _close_hwp(hwp) -> None:
        """
        HWP COM 객체 해제.
        Quit() / Clear() 모두 DRM 환경에서 팝업 유발 → 사용 안 함.
        del로 COM 참조를 끊으면 HWP 프로세스가 자동 종료됨.
        """
        try:
            del hwp
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 텍스트 추출
    # ------------------------------------------------------------------

    def _extract_sections(self, hwp) -> list[Section]:
        """
        GetTextFile("TEXT", "") 로 문서 전체 텍스트를 한 번에 추출.

        InitScan/GetText 루프는 선택 범위·컨트롤 처리 순서에 민감해서
        환경에 따라 무한루프·빈 결과가 발생함.
        GetTextFile은 HWP가 내부적으로 변환해 주므로 가장 안정적.

        반환 형식 "TEXT" → 줄바꿈(\r\n 또는 \n)으로 단락이 구분된 평문.
        """
        main_section = Section(index=0)

        raw = ""
        try:
            raw = hwp.GetTextFile("TEXT", "")
        except Exception:
            pass

        if raw:
            import html as _html
            raw = _html.unescape(raw)
            raw = self._clean_hwp_text(raw)
            lines = raw.replace("\r\n", "\n").replace("\r", "\n").split("\n")
            for line in lines:
                para = Paragraph(runs=[TextRun(line)] if line.strip() else [])
                main_section.blocks.append(para)
            return [main_section]

        # fallback: SelectAll + GetText
        try:
            hwp.Run("SelectAll")
            hwp.InitScan(option=None, Range=0xff,
                         spara=None, spos=None, epara=None, epos=None)
            state, text = hwp.GetText()
            hwp.ReleaseScan()
            hwp.Run("Cancel")
            if text:
                lines = text.replace("\r", "\n").split("\n")
                for line in lines:
                    para = Paragraph(runs=[TextRun(line)] if line.strip() else [])
                    main_section.blocks.append(para)
        except Exception:
            pass

        return [main_section]

    @staticmethod
    def _flush_para(buf: list[str], section: Section) -> None:
        """버퍼의 텍스트를 Paragraph로 변환하여 섹션에 추가, 버퍼 초기화"""
        text = "".join(buf).strip()
        buf.clear()
        para = Paragraph(runs=[TextRun(text)] if text else [])
        section.blocks.append(para)

    @staticmethod
    def _extract_metadata(hwp, step=None) -> dict[str, str]:
        """
        DRM 환경에서 XHwpDocuments / Summary 접근이
        '보안정책상 사용할 수 없는 기능' 팝업을 유발함.
        메타데이터 추출을 완전히 건너뜀.
        """
        return {}

    # ------------------------------------------------------------------
    # HWP 특수문자 정제
    # ------------------------------------------------------------------

    # GetTextFile("TEXT") 가 변환하는 HWP 내부 특수문자 → 대체 문자열 매핑
    # U+25E6 ◦ : 글머리 기호(○계열)   U+2022 • : 글머리 기호(●계열)
    # U+0002    : HWP 내부 컨트롤 코드  U+0005    : 필드 코드
    # U+000B    : 수직탭(셀 구분)       U+001C    : 파일 구분자
    _HWP_CHAR_MAP: dict[str, str] = {
        "\u25e6": "◦",  # ◦  흰 글머리표 유지
        "\u2022": "•",  # •  검은 글머리표 유지
        "\u25cf": "●",  # ●  검은 원 유지
        "\u25cb": "○",  # ○  흰 원 유지
        "\u0002": "",   # STX 컨트롤 코드
        "\u0005": "",   # 필드 코드
        "\u000b": "\n", # 수직탭 → 줄바꿈
        "\u001c": "",   # 파일 구분자
        "\u001d": "",   # 그룹 구분자
        "\u001e": "",   # 레코드 구분자
        "\u001f": "",   # 단위 구분자
        "\xa0":   " ",  # non-breaking space → 일반 공백
    }

    @classmethod
    def _clean_hwp_text(cls, text: str) -> str:
        """HWP GetTextFile 특수문자 정제"""
        for src, dst in cls._HWP_CHAR_MAP.items():
            text = text.replace(src, dst)
        # C0 제어문자(탭·줄바꿈 제외) 제거
        text = re.sub(r"[\x00-\x08\x0c\x0e-\x1b]", "", text)
        return text


# ---------------------------------------------------------------------------
# HWPX 아카이브 리더 (하위 호환 유지용 래퍼)
# ---------------------------------------------------------------------------

class HWPXArchiveReader:
    """
    [하위 호환] ZipDocReader의 저수준 ZIP 접근 래퍼.
    SectionParser와 분리된 파이프라인이 필요할 때 직접 사용.
    """

    def __init__(self, path: Path) -> None:
        self.path = path
        self._zf: Optional[zipfile.ZipFile] = None

    def __enter__(self) -> "HWPXArchiveReader":
        if ZipDocReader._is_drm_protected(self.path):
            raise HWPXDrmError(
                f"{self.path.name}: DRM 보호 파일 — COM 백엔드를 사용하세요."
            )
        try:
            self._zf = zipfile.ZipFile(self.path, "r")
        except zipfile.BadZipFile as exc:
            raise HWPXDrmError(
                f"{self.path.name}은 ZIP으로 열 수 없습니다 (DRM 가능성)."
            ) from exc
        return self

    def __exit__(self, *args: object) -> None:
        if self._zf:
            self._zf.close()

    @property
    def namelist(self) -> list[str]:
        assert self._zf, "컨텍스트 매니저 내에서 사용하세요"
        return self._zf.namelist()

    def read(self, name: str) -> bytes:
        assert self._zf
        return self._zf.read(name)

    def section_entries(self) -> list[tuple[int, str]]:
        return ZipDocReader._section_entries(self._zf)  # type: ignore

    def read_metadata(self) -> dict[str, str]:
        return ZipDocReader._read_metadata(self._zf)  # type: ignore


# ---------------------------------------------------------------------------
# ParserFactory (GPT 버전 참조)
# ---------------------------------------------------------------------------

class ParserFactory:
    """
    파일 확장자 기반으로 HWPXParser를 생성하는 팩토리.

    현재 지원:
      .hwpx → HWPXParser (ZIP 또는 COM 백엔드)
      .hwp  → HWPXParser (COM 백엔드 전용, ZIP은 구 바이너리 형식)

    사용 예::

        parser = ParserFactory.create("문서.hwpx")
        text = parser.parse_text("문서.hwpx")

        # DRM 환경
        parser = ParserFactory.create("문서.hwpx", backend="com")
    """

    @staticmethod
    def create(
        file_path: str | Path,
        backend: str = "auto",
        com_visible: bool = False,
    ) -> "HWPXParser":
        path = Path(file_path)
        suffix = path.suffix.lower()

        if suffix == ".hwpx":
            return HWPXParser(backend=backend, com_visible=com_visible)
        if suffix == ".hwp":
            # 구 바이너리 .hwp는 ZIP 파싱 불가 → COM 강제
            if backend == "auto":
                LOGGER.debug(".hwp 파일 감지: COM 백엔드로 자동 전환")
                return HWPXParser(backend="com", com_visible=com_visible)
            return HWPXParser(backend=backend, com_visible=com_visible)

        raise HWPXFormatError(
            f"지원하지 않는 파일 형식: {path.suffix}. "
            ".hwpx 또는 .hwp 파일을 사용하세요."
        )


# ---------------------------------------------------------------------------
# CLI 헬퍼 함수 (GPT 버전 참조)
# ---------------------------------------------------------------------------

def build_output_path(
    input_path: Path,
    output_path: Optional[str | Path] = None,
) -> Path:
    """출력 파일 경로 결정. 미지정 시 입력 파일과 같은 위치에 .txt로 저장."""
    if output_path is not None:
        return Path(output_path)
    return input_path.with_suffix(".txt")


def configure_logging(verbose: bool = False) -> None:
    """로깅 설정. verbose=True이면 DEBUG, 아니면 INFO."""
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(levelname)s [%(name)s]: %(message)s",
    )


# ---------------------------------------------------------------------------
# 메인 파서 (Facade)
# ---------------------------------------------------------------------------

class HWPXParser:
    """
    HWPX / HWP 파일 파서 — 주요 진입점(Facade).

    backend 옵션
    ------------
    "auto" (기본값)
        ZIP 직접 파싱을 시도하고, DRM 보호 파일이면 자동으로
        COM 백엔드로 전환합니다.
    "zip"
        ZIP 직접 파싱만 사용합니다. DRM 파일에서는 HWPXDrmError 발생.
    "com"
        COM(win32com) 백엔드만 사용합니다.
        Windows + 한글(HWP) 설치가 필수입니다.

    사용 예::

        # 자동 (DRM 파일 포함 처리)
        parser = HWPXParser()
        doc = parser.parse("문서.hwpx")
        print(doc.extract_text())

        # COM 강제 (DRM 확실한 환경)
        parser = HWPXParser(backend="com")
        doc = parser.parse("drm_문서.hwpx")

        # HWP 창 표시 (일부 DRM 환경 필요)
        parser = HWPXParser(backend="com", com_visible=True)

        # 커스텀 블록 파서 등록
        class MyParser(BaseNodeParser): ...
        parser = HWPXParser()
        parser.section_parser.register(MyParser())
    """

    def __init__(
        self,
        backend: str = "auto",
        com_visible: bool = False,
    ) -> None:
        """
        Args:
            backend:     "auto" | "zip" | "com"
            com_visible: COM 백엔드에서 HWP 창 표시 여부 (기본 False=숨김)
        """
        try:
            self._backend = Backend[backend.upper()]
        except KeyError:
            raise ValueError(
                f"backend={backend!r} 는 유효하지 않습니다. "
                f"'auto', 'zip', 'com' 중 하나를 선택하세요."
            )
        self._com_visible = com_visible
        self.section_parser = SectionParser()

        # 백엔드 인스턴스 (지연 생성)
        self._zip_reader = ZipDocReader()
        self._com_reader: Optional[ComDocReader] = None

    # ------------------------------------------------------------------
    # 공개 API
    # ------------------------------------------------------------------

    def parse(self, path: str | Path) -> HWPXDocument:
        """
        파일을 파싱하여 HWPXDocument 반환.

        Args:
            path: .hwpx 또는 .hwp 파일 경로

        Returns:
            HWPXDocument 인스턴스

        Raises:
            FileNotFoundError: 파일이 없을 때
            HWPXDrmError:      DRM 파일에서 ZIP 백엔드 강제 시
            HWPXComError:      COM 초기화 실패 (HWP 미설치 등)
            HWPXFormatError:   유효하지 않은 형식
            HWPXParseError:    XML 파싱 실패
        """
        path = Path(path)
        if not path.exists():
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {path}")
        if not path.is_file():
            raise HWPXFormatError(f"파일이 아닙니다: {path}")

        if self._backend == Backend.ZIP:
            return self._zip_reader.read_document(path, self.section_parser)

        if self._backend == Backend.COM:
            return self._get_com_reader().read_document(path, self.section_parser)

        # --- AUTO: ZIP 시도 → DRM 감지 시 COM 전환 ---
        try:
            return self._zip_reader.read_document(path, self.section_parser)
        except HWPXDrmError as drm_err:
            LOGGER.info(
                "%s: DRM 보호 파일 감지 — COM 백엔드로 전환합니다. (%s)",
                path.name, drm_err,
            )
            # DRM 파일: COM 백엔드로 재시도
            try:
                com_reader = self._get_com_reader()
            except HWPXComError as com_err:
                # COM도 사용 불가 → 두 오류를 모두 명시
                raise HWPXComError(
                    f"{path.name}: DRM 파일이지만 COM 백엔드를 사용할 수 없습니다.\n"
                    f"  DRM 오류: {drm_err}\n"
                    f"  COM 오류: {com_err}\n"
                    "  Windows 환경에서 한글(HWP)을 설치하고 pywin32를 설치하세요."
                ) from com_err
            return com_reader.read_document(path, self.section_parser)

    def parse_text(
        self,
        path: str | Path,
        options: Optional[TextExtractionOptions] = None,
        *,
        include_tables: bool = True,
        skip_empty: bool = True,
    ) -> str:
        """
        편의 메서드 — 파싱 후 바로 텍스트 반환.

        Args:
            path:    .hwpx / .hwp 파일 경로
            options: TextExtractionOptions (None이면 include_tables/skip_empty 적용)
        """
        doc = self.parse(path)
        return doc.extract_text(
            options,
            include_tables=include_tables,
            skip_empty=skip_empty,
        )

    @property
    def backend_name(self) -> str:
        """현재 설정된 백엔드 이름"""
        return self._backend.name.lower()

    # ------------------------------------------------------------------
    # 내부
    # ------------------------------------------------------------------

    def _get_com_reader(self) -> ComDocReader:
        if self._com_reader is None:
            self._com_reader = ComDocReader(visible=self._com_visible)
        return self._com_reader


# ---------------------------------------------------------------------------
# 대화형 CLI  (Interactive CLI)
# ---------------------------------------------------------------------------

# ── 출력 헬퍼 ───────────────────────────────────────────────────────────────

_W = 60  # 구분선 너비

def _hr(char: str = "─") -> None:
    print(char * _W)

def _section(title: str) -> None:
    print()
    _hr("─")
    print(f"  {title}")
    _hr("─")

def _ok(msg: str) -> None:
    print(f"  ✔  {msg}")

def _err(msg: str) -> None:
    print(f"  ✘  {msg}")

def _info(msg: str) -> None:
    print(f"     {msg}")


def _ask(prompt: str, default: str = "") -> str:
    """input() 래퍼. 기본값 표시 및 빈 입력 처리."""
    hint = f" [기본: {default}]" if default else ""
    raw = input(f"  → {prompt}{hint}: ").strip()
    return raw if raw else default


def _ask_yn(prompt: str, default: bool = True) -> bool:
    """Y/N 질문. 기본값에 따라 대문자 표시."""
    hint = "Y/n" if default else "y/N"
    raw = input(f"  → {prompt} ({hint}): ").strip().lower()
    if raw in ("y", "yes"):
        return True
    if raw in ("n", "no"):
        return False
    return default


def _ask_choice(prompt: str, choices: list[tuple[str, str]], default: int = 1) -> str:
    """
    번호 선택 메뉴.
    choices: [(값, 설명), ...]
    반환:    선택된 값 문자열
    """
    for i, (_, desc) in enumerate(choices, 1):
        marker = " ◀ 기본값" if i == default else ""
        print(f"     {i}) {desc}{marker}")
    while True:
        raw = input(f"  → {prompt} (번호 입력, Enter=기본값): ").strip()
        if raw == "":
            return choices[default - 1][0]
        if raw.isdigit() and 1 <= int(raw) <= len(choices):
            return choices[int(raw) - 1][0]
        _err(f"1~{len(choices)} 사이의 번호를 입력하세요.")


# ── 단계별 함수 ─────────────────────────────────────────────────────────────

def _step_select_files() -> list[Path]:
    """[1단계] tkinter 파일 대화창으로 파일 선택 (복수 가능)."""
    _section("1단계 | 파일 선택")

    try:
        import tkinter as tk
        from tkinter import filedialog
        _tk_available = True
    except ImportError:
        _tk_available = False

    paths: list[Path] = []

    if _tk_available:
        print("  파일 선택 대화창을 열겠습니다.")
        ans = _ask_yn("대화창 열기", default=True)
        if ans:
            root = tk.Tk()
            root.withdraw()          # 메인 창 숨기기
            root.attributes("-topmost", True)
            selected = filedialog.askopenfilenames(
                title="HWPX / HWP 파일 선택 (여러 파일 선택 가능)",
                filetypes=[
                    ("한글 문서", "*.hwpx *.hwp"),
                    ("HWPX 파일", "*.hwpx"),
                    ("HWP 파일",  "*.hwp"),
                    ("모든 파일", "*.*"),
                ],
            )
            root.destroy()
            paths = [Path(p) for p in selected]

    # 대화창 미사용 또는 tkinter 없음 → 경로 직접 입력
    if not paths:
        if not _tk_available:
            _err("tkinter를 사용할 수 없습니다. 경로를 직접 입력하세요.")
        else:
            _info("파일이 선택되지 않았습니다. 경로를 직접 입력합니다.")

        print("  (여러 파일: 쉼표로 구분, 예: C:\\a.hwpx, C:\\b.hwpx)")
        raw = _ask("파일 경로 입력")
        if not raw:
            return []
        paths = [Path(p.strip()) for p in raw.split(",") if p.strip()]

    if not paths:
        _err("파일이 선택되지 않았습니다.")
        return []

    print()
    _ok(f"{len(paths)}개 파일 선택됨:")
    for p in paths:
        _info(str(p))

    return paths


def _step_backend() -> tuple[str, bool]:
    """[2단계] 백엔드 및 COM 가시성 선택."""
    _section("2단계 | 파싱 백엔드 선택")

    backend = _ask_choice(
        "백엔드",
        choices=[
            ("auto", "auto  — 자동 (비DRM: ZIP, DRM: COM 자동 전환)"),
            ("zip",  "zip   — ZIP 직접 파싱만 (비DRM 전용, 빠름)"),
            ("com",  "com   — HWP COM 자동화 (DRM 파일 / .hwp)"),
        ],
        default=1,
    )

    com_visible = False
    if backend in ("com", "auto"):
        print()
        print("  COM 백엔드 옵션:")
        com_visible = _ask_yn("HWP 창 화면에 표시 (일부 DRM 환경 필요)", default=False)

    _ok(f"백엔드: {backend}" + (" / HWP 창 표시" if com_visible else ""))
    return backend, com_visible


def _step_options() -> TextExtractionOptions:
    """[3단계] 텍스트 추출 옵션 설정."""
    _section("3단계 | 추출 옵션")

    include_tables    = _ask_yn("표(Table) 내용 포함",   default=True)
    preserve_blanks   = _ask_yn("빈 줄 유지",             default=True)
    normalize_ws      = _ask_yn("내부 공백 정규화",       default=True)

    opt = TextExtractionOptions(
        include_tables=include_tables,
        preserve_blank_lines=preserve_blanks,
        normalize_whitespace=normalize_ws,
        strip_lines=True,
        section_separator="\n\n",
    )

    print()
    _ok("옵션 확정:")
    _info(f"표 포함={include_tables}  빈줄유지={preserve_blanks}  공백정규화={normalize_ws}")
    return opt


def _step_parse(
    paths: list[Path],
    backend: str,
    com_visible: bool,
    options: TextExtractionOptions,
) -> list[tuple[Path, HWPXDocument]]:
    """[4단계] 파싱 실행."""
    _section("4단계 | 파싱 실행")

    parser = ParserFactory.create(paths[0], backend=backend, com_visible=com_visible)
    results: list[tuple[Path, HWPXDocument]] = []

    for path in paths:
        print(f"  처리 중: {path.name} ...", end="", flush=True)
        try:
            # .hwp는 팩토리에서 com 강제이므로 그대로 parse
            doc = parser.parse(path)
            results.append((path, doc))
            n_para  = sum(len(s.paragraphs) for s in doc.sections)
            n_tbl   = sum(len(s.tables)     for s in doc.sections)
            print(f" 완료")
            _info(
                f"섹션 {len(doc.sections)}개 / "
                f"단락 {n_para}개 / "
                f"표 {n_tbl}개"
                + (f" / 백엔드: {parser.backend_name}" )
            )
            if doc.metadata:
                for k, v in doc.metadata.items():
                    _info(f"  {k}: {v}")
        except HWPXError as exc:
            print(f" 실패")
            _err(str(exc))
        except FileNotFoundError as exc:
            print(f" 실패")
            _err(str(exc))

    return results


def _step_output(results: list[tuple[Path, HWPXDocument]], options: TextExtractionOptions) -> None:
    """[5단계] 결과 출력 방식 선택 및 실행."""
    if not results:
        _err("저장할 결과가 없습니다.")
        return

    _section("5단계 | 결과 저장")

    mode = _ask_choice(
        "출력 방식",
        choices=[
            ("screen",  "화면에 출력 (stdout)"),
            ("auto",    "입력 파일 옆에 자동으로 .txt 저장"),
            ("manual",  "저장 경로 직접 입력"),
        ],
        default=2,
    )

    if mode == "screen":
        for path, doc in results:
            if len(results) > 1:
                _hr("=")
                print(f"  ▶ {path.name}")
                _hr("=")
            print(doc.extract_text(options))

    elif mode == "auto":
        for path, doc in results:
            out = build_output_path(path)
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(doc.extract_text(options), encoding="utf-8")
            _ok(f"저장 완료: {out}")

    elif mode == "manual":
        if len(results) == 1:
            path, doc = results[0]
            default_out = str(build_output_path(path))
            raw = _ask("저장 경로", default=default_out)
            out = Path(raw)
            out.parent.mkdir(parents=True, exist_ok=True)
            out.write_text(doc.extract_text(options), encoding="utf-8")
            _ok(f"저장 완료: {out}")
        else:
            # 복수 파일: 저장 디렉터리 지정 후 각각 .txt로 저장
            print("  복수 파일 → 저장 폴더를 지정하거나 Enter로 각 파일 옆에 저장합니다.")
            raw_dir = _ask("저장 폴더 (비워두면 각 파일 위치)").strip()
            for path, doc in results:
                if raw_dir:
                    out = Path(raw_dir) / (path.stem + ".txt")
                else:
                    out = build_output_path(path)
                out.parent.mkdir(parents=True, exist_ok=True)
                out.write_text(doc.extract_text(options), encoding="utf-8")
                _ok(f"저장 완료: {out}")


# ── 메인 루프 ────────────────────────────────────────────────────────────────

def main() -> int:
    configure_logging(verbose=False)

    print()
    _hr("═")
    print("  HWPParser  |  HWPX / HWP 텍스트 추출기")
    print("  표준 라이브러리 + COM(선택) 기반 / DRM 지원")
    _hr("═")

    while True:
        # ── 1단계: 파일 선택 ──────────────────────────────────────────
        paths = _step_select_files()
        if not paths:
            print()
            _err("파일이 없으므로 종료합니다.")
            break

        # ── 2단계: 백엔드 ────────────────────────────────────────────
        backend, com_visible = _step_backend()

        # ── 3단계: 추출 옵션 ─────────────────────────────────────────
        options = _step_options()

        # ── 4단계: 파싱 ──────────────────────────────────────────────
        results = _step_parse(paths, backend, com_visible, options)

        # ── 5단계: 출력 ──────────────────────────────────────────────
        if results:
            _step_output(results, options)

        # ── 6단계: 계속 여부 ─────────────────────────────────────────
        _section("6단계 | 계속 여부")
        again = _ask_yn("다른 파일을 처리하시겠습니까?", default=False)
        if not again:
            print()
            _hr("═")
            print("  종료합니다.")
            _hr("═")
            print()
            break

    return 0


if __name__ == "__main__":
    import sys
    sys.exit(main())