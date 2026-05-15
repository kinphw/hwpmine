"""
Step 3 — 문서 검색 GUI
=======================
MariaDB에 적재된 문서(HWP·PDF 통합)를 키워드로 검색하고 클릭하면 파일을 엽니다.
HWP 와 PDF 는 적재 단계만 분리돼 있고 검색은 단일 테이블에서 통합 수행.

단독 실행:
  python search_gui.py
"""

import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

try:
    import pymysql
except ImportError:
    raise SystemExit("pymysql 필요: pip install pymysql")

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    _DND_AVAILABLE = True
except ImportError:
    TkinterDnD = None  # type: ignore[assignment]
    DND_FILES = "DND_Files"
    _DND_AVAILABLE = False

from . import config
from .about import show_about
from .icon import make_app_icon

PAGE_SIZE = 200


# ═══════════════════════════════════════════════════════════════
# 클립보드 파일 복사 — Ctrl+C → Explorer 에서 Ctrl+V
# ═══════════════════════════════════════════════════════════════
#
# 운영망 DRM 이 프로세스 basename ('python.exe') 기준으로 drag-and-drop
# (DoDragDrop) 을 막는 환경에서, 같은 기능을 클립보드 경로로 제공한다.
# DRM 솔루션에 따라 OleSetClipboard 도 함께 후킹돼 막힐 수 있는데, 그건
# 코드 차원에서 더 우회할 길이 없어 솔루션 정책에 달려있다.
#
# Tkinter 기본 clipboard API 는 텍스트만 다루므로, Windows 의 CF_HDROP
# 포맷(DROPFILES 구조체 + UTF-16 paths) 을 직접 만들어 win32clipboard 로
# 올린다. Explorer 는 이 포맷을 인식해 Ctrl+V 시 파일 복사로 처리한다.

def _copy_files_to_clipboard(paths: list[str]) -> None:
    """파일 경로 리스트를 Explorer 호환 CF_HDROP 포맷으로 클립보드에 올림."""
    import struct
    import win32clipboard
    import win32con

    # null 종료된 UTF-16 문자열 + 최종 종료(double null)
    paths_str = "\0".join(paths) + "\0\0"
    paths_bytes = paths_str.encode("utf-16-le")

    # DROPFILES 구조체: DWORD pFiles=20, POINT pt=(0,0), BOOL fNC=0, BOOL fWide=1
    # (헤더 크기 20바이트 = 4 + 4 + 4 + 4 + 4)
    dropfiles = struct.pack("<Iiiii", 20, 0, 0, 0, 1)
    data = dropfiles + paths_bytes

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_HDROP, data)
    finally:
        win32clipboard.CloseClipboard()


# ═══════════════════════════════════════════════════════════════
# DB
# ═══════════════════════════════════════════════════════════════

def get_conn():
    return pymysql.connect(**config.get_db_config())


def _prepare_keywords(keyword: str, mode: str) -> list[str]:
    if mode == "phrase":
        return [keyword] if keyword else []
    return keyword.split() or []


# 키워드 양옆에 잡을 문자 수 — 비대칭. 미리보기 셀(width=500) 이 좁아서 키워드가
# 중앙에 있으면 잘려보이므로, 키워드를 셀 좌측 가까이 끌어오기 위해 left << right.
SNIPPET_LEFT  = 20
SNIPPET_RIGHT = 220


def _extract_snippet(body: str, keywords: list[str],
                     left: int = SNIPPET_LEFT, right: int = SNIPPET_RIGHT) -> str:
    """본문에서 가장 먼저 나오는 키워드 매치 위치 기준 앞 `left` / 뒤 `right` 자를 잘라 반환.
    키워드 없거나 매칭 위치를 못 찾으면 본문 앞부분으로 폴백.
    SQL 단에서 LIKE 로 매칭된 행만 들어오지만 LEFT(5000) 절단으로 매치가 잘려나간
    경우도 있어 폴백이 필요하다."""
    if not body:
        return ""
    body = body.replace("\r", "").replace("\n", " ")
    fallback_len = left + right + 30  # 폴백 시 표시할 앞부분 길이

    if not keywords:
        return body[:fallback_len].strip()

    # 모든 키워드의 첫 매치 중 가장 빠른 위치를 채택 (대소문자 무시).
    body_lower = body.lower()
    best_pos, best_kw_len = -1, 0
    for kw in keywords:
        if not kw:
            continue
        pos = body_lower.find(kw.lower())
        if pos < 0:
            continue
        if best_pos < 0 or pos < best_pos:
            best_pos, best_kw_len = pos, len(kw)

    if best_pos < 0:
        return body[:fallback_len].strip()

    start = max(0, best_pos - left)
    end   = min(len(body), best_pos + best_kw_len + right)
    snippet = body[start:end].strip()
    if start > 0:
        snippet = "…" + snippet
    if end < len(body):
        snippet = snippet + "…"
    return snippet


def _build_where(keywords: list[str], target: str, mode: str):
    conds, params = [], []
    for kw in keywords:
        like = f"%{kw}%"
        if target == "title":
            conds.append("filename LIKE %s")
            params.append(like)
        elif target == "body":
            conds.append("body_text LIKE %s")
            params.append(like)
        else:
            conds.append("(filename LIKE %s OR body_text LIKE %s)")
            params.extend([like, like])
    joiner = " AND " if mode in ("and", "phrase") else " OR "
    return joiner.join(conds), params


def _compose_where(keyword: str, target: str, mode: str, include_excluded: bool,
                   id_min=None, id_max=None):
    keywords  = _prepare_keywords(keyword, mode)
    conds, params = [], []

    if not include_excluded:
        conds.append("body_text IS NOT NULL AND body_text != ''")
    if keywords:
        kw_where, kw_params = _build_where(keywords, target, mode)
        conds.append(f"({kw_where})")
        params.extend(kw_params)
    if id_min is not None:
        conds.append("id >= %s")
        params.append(id_min)
    if id_max is not None:
        conds.append("id <= %s")
        params.append(id_max)

    where_sql = (" WHERE " + " AND ".join(conds)) if conds else ""
    return where_sql, params


def search(keyword: str, target: str, mode: str = "and",
           limit: int = PAGE_SIZE, offset: int = 0, include_excluded: bool = False,
           id_min=None, id_max=None):
    where_sql, params = _compose_where(keyword, target, mode, include_excluded, id_min, id_max)
    # body_text 를 5000자까지 가져와서 클라이언트에서 키워드 인근으로 스니펫을 잘라낸다.
    # 200건 × 5KB ≈ 1MB. 5000자 내에 키워드가 없는 케이스는 _extract_snippet 이 앞부분 폴백.
    sql = f"""
        SELECT id, directory, filename, LEFT(body_text, 5000)
        FROM `{config.DB_TABLE}`
        {where_sql}
        ORDER BY id
        LIMIT %s OFFSET %s
    """
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, (*params, limit, offset))
            return cur.fetchall()
    finally:
        conn.close()


def count_results(keyword: str, target: str, mode: str = "and", include_excluded: bool = False,
                  id_min=None, id_max=None) -> int:
    where_sql, params = _compose_where(keyword, target, mode, include_excluded, id_min, id_max)
    sql = f"SELECT COUNT(*) FROM `{config.DB_TABLE}`{where_sql}"
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, params)
            return cur.fetchone()[0]
    finally:
        conn.close()


def nullify_body_text(ids: list) -> int:
    """레코드는 유지하되 body_text를 NULL 처리해 검색에서 제외한다."""
    if not ids:
        return 0
    placeholders = ", ".join(["%s"] * len(ids))
    sql = f"UPDATE `{config.DB_TABLE}` SET body_text = NULL WHERE id IN ({placeholders})"
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, ids)
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


def delete_rows(ids: list) -> int:
    """레코드 자체를 DB에서 완전히 삭제한다."""
    if not ids:
        return 0
    placeholders = ", ".join(["%s"] * len(ids))
    sql = f"DELETE FROM `{config.DB_TABLE}` WHERE id IN ({placeholders})"
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            cur.execute(sql, ids)
        conn.commit()
        return cur.rowcount
    finally:
        conn.close()


# ═══════════════════════════════════════════════════════════════
# GUI
# ═══════════════════════════════════════════════════════════════

class App:
    def __init__(self, master: tk.Misc):
        # master 가 Tk/Toplevel 이면 단독 창 모드, 그 외(Notebook 탭의 Frame 등)는 임베드 모드.
        self.root = master
        self._standalone = isinstance(master, (tk.Tk, tk.Toplevel))

        if self._standalone:
            master.title("문서 검색기")
            master.geometry("1100x750")
            master.minsize(800, 500)

            # 아이콘은 PhotoImage GC 방지를 위해 인스턴스 속성으로 보관
            try:
                self._app_icon = make_app_icon(master)
                master.iconphoto(True, self._app_icon)
            except tk.TclError:
                self._app_icon = None

        self.results: list = []
        self._full_data: dict = {}
        self._excluded_ids: set = set()
        self._sel_anchor  = None
        self._tooltip     = None
        self._offset      = 0
        self._total       = 0
        self._last_kw              = ""
        self._last_target          = "both"
        self._last_mode            = "and"
        self._last_include_excluded = False
        self._last_id_min          = None
        self._last_id_max          = None

        self._build_ui()
        self._bind_events()

    def _build_ui(self):
        # ── 상단: 검색 컨트롤 ────────────────────────────────────
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill=tk.X)

        row1 = ttk.Frame(top)
        row1.pack(fill=tk.X)

        ttk.Label(row1, text="키워드:").pack(side=tk.LEFT)
        self.entry = ttk.Entry(row1, width=40, font=("맑은 고딕", 11))
        self.entry.pack(side=tk.LEFT, padx=(5, 10))

        self.btn = ttk.Button(row1, text="검색", command=self._on_search)
        self.btn.pack(side=tk.LEFT)

        ttk.Button(row1, text="?", width=3,
                   command=lambda: show_about(self.root)).pack(side=tk.RIGHT, padx=(8, 0))

        self.status = ttk.Label(row1, text="", foreground="gray")
        self.status.pack(side=tk.RIGHT)

        row2 = ttk.Frame(top)
        row2.pack(fill=tk.X, pady=(5, 0))

        ttk.Label(row2, text="검색 대상:").pack(side=tk.LEFT)
        self.target_var = tk.StringVar(value="both")
        ttk.Radiobutton(row2, text="제목+본문", variable=self.target_var, value="both").pack(side=tk.LEFT)
        ttk.Radiobutton(row2, text="제목만",   variable=self.target_var, value="title").pack(side=tk.LEFT)
        ttk.Radiobutton(row2, text="본문만",   variable=self.target_var, value="body").pack(side=tk.LEFT)

        ttk.Separator(row2, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=12)

        ttk.Label(row2, text="검색 방식:").pack(side=tk.LEFT)
        self.mode_var = tk.StringVar(value="and")
        ttk.Radiobutton(row2, text="AND (공백=모두 포함)", variable=self.mode_var, value="and").pack(side=tk.LEFT)
        ttk.Radiobutton(row2, text="OR (공백=하나라도)",   variable=self.mode_var, value="or").pack(side=tk.LEFT)
        ttk.Radiobutton(row2, text="전체 문자열",          variable=self.mode_var, value="phrase").pack(side=tk.LEFT)

        ttk.Separator(row2, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=12)
        self.include_excluded_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row2, text="제외 항목 포함", variable=self.include_excluded_var,
            command=self._on_search,
        ).pack(side=tk.LEFT)

        # ── ID 범위 필터 (우측) ─────────────────────────────────
        self.id_max_var   = tk.StringVar()
        self.id_max_entry = ttk.Entry(row2, width=9, textvariable=self.id_max_var)
        self.id_max_entry.pack(side=tk.RIGHT, padx=(2, 0))
        ttk.Label(row2, text="ID ≤").pack(side=tk.RIGHT, padx=(6, 2))

        self.id_min_var   = tk.StringVar()
        self.id_min_entry = ttk.Entry(row2, width=9, textvariable=self.id_min_var)
        self.id_min_entry.pack(side=tk.RIGHT, padx=(2, 0))
        ttk.Label(row2, text="ID ≥").pack(side=tk.RIGHT, padx=(6, 2))

        self.id_filter_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            row2, text="ID 범위", variable=self.id_filter_var,
            command=self._on_toggle_id_filter,
        ).pack(side=tk.RIGHT)
        self._on_toggle_id_filter()  # 초기 비활성

        # ── 중간: Treeview ────────────────────────────────────────
        mid = ttk.Frame(self.root)
        mid.pack(fill=tk.BOTH, expand=True, padx=10)

        cols = ("id", "directory", "filename", "preview")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings", selectmode="extended")

        self.tree.heading("id",        text="ID")
        self.tree.heading("directory", text="폴더")
        self.tree.heading("filename",  text="파일명")
        self.tree.heading("preview",   text="내용 미리보기")

        self.tree.column("id",        width=50,  minwidth=40,  stretch=False)
        self.tree.column("directory", width=300, minwidth=100)
        self.tree.column("filename",  width=280, minwidth=100)
        self.tree.column("preview",   width=500, minwidth=100)

        vsb = ttk.Scrollbar(mid, orient=tk.VERTICAL,   command=self.tree.yview)
        hsb = ttk.Scrollbar(mid, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.tag_configure("excluded", foreground="#999999")

        if _DND_AVAILABLE:
            self.tree.drag_source_register(1, DND_FILES)
            self.tree.dnd_bind("<<DragInitCmd>>", self._on_drag_init)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        mid.grid_rowconfigure(0, weight=1)
        mid.grid_columnconfigure(0, weight=1)

        # ── 로그창 ────────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self.root, text="로그", padding=4)
        log_frame.pack(fill=tk.X, padx=10, pady=(4, 0))

        self.log_text = scrolledtext.ScrolledText(
            log_frame, height=5, state="disabled",
            font=("Consolas", 8), wrap="word",
            bg="#1e1e1e", fg="#d4d4d4",
        )
        self.log_text.pack(fill=tk.X)

        # ── 하단: 버튼 ───────────────────────────────────────────
        bot = ttk.Frame(self.root, padding=10)
        bot.pack(fill=tk.X)

        hint = "더블클릭: 파일 열기 | 드래그/Ctrl+C: 탐색기로 복사 | Del: 검색에서 제외"
        if not _DND_AVAILABLE:
            hint = "더블클릭: 파일 열기 | Ctrl+C: 탐색기로 복사 | Del: 검색에서 제외"
        self.info_label = ttk.Label(
            bot,
            text=hint,
            foreground="gray", font=("맑은 고딕", 9),
        )
        self._info_default = hint
        self.info_label.pack(side=tk.LEFT)

        self.open_btn = ttk.Button(bot, text="파일 열기",  command=self._on_open,   state=tk.DISABLED)
        self.open_btn.pack(side=tk.RIGHT)

        self.open_path_btn = ttk.Button(bot, text="경로 열기", command=self._on_open_path, state=tk.DISABLED)
        self.open_path_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self.del_btn  = ttk.Button(bot, text="제외 (Del)", command=self._on_delete, state=tk.DISABLED)
        self.del_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self.all_btn  = ttk.Button(bot, text="전체 조회",  command=self._on_load_all, state=tk.DISABLED)
        self.all_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self.more_btn = ttk.Button(bot, text=f"더보기 (+{PAGE_SIZE})", command=self._on_more, state=tk.DISABLED)
        self.more_btn.pack(side=tk.RIGHT, padx=(0, 5))

        # 우클릭 컨텍스트 메뉴 — 항목은 매번 _on_right_click 에서 선택 상태 +
        # 클릭한 컬럼에 맞춰 동적으로 재빌드한다.
        self.context_menu = tk.Menu(self.root, tearoff=0)

    def _bind_events(self):
        self.entry.bind("<Return>",            lambda e: self._on_search())
        self.id_min_entry.bind("<Return>",     lambda e: self._on_search())
        self.id_max_entry.bind("<Return>",     lambda e: self._on_search())
        self.tree.bind("<Double-1>",           lambda e: self._on_open())
        self.tree.bind("<<TreeviewSelect>>",   self._on_select)
        self.tree.bind("<Motion>",             self._on_hover)
        self.tree.bind("<Leave>",              self._hide_tooltip)
        self.tree.bind("<Delete>",             lambda e: self._on_delete())
        self.tree.bind("<Shift-Up>",           lambda e: self._extend_selection(-1))
        self.tree.bind("<Shift-Down>",         lambda e: self._extend_selection(1))
        self.tree.bind("<Up>",                 self._reset_anchor_on_move)
        self.tree.bind("<Down>",               self._reset_anchor_on_move)
        self.tree.bind("<Button-1>",           self._reset_anchor_on_move)
        # Ctrl+C / Ctrl+Insert — 선택 파일들을 클립보드에 올려 Explorer 등에서
        # Ctrl+V 로 붙여넣기 복사. DRM 이 drag 를 막는 환경의 우회용.
        # Ctrl-C 와 Control-c 둘 다 바인딩 — IME/CapsLock 상태에 따라 다르게 옴.
        self.tree.bind("<Control-c>",          lambda e: self._on_copy_files())
        self.tree.bind("<Control-C>",          lambda e: self._on_copy_files())
        self.tree.bind("<Control-Insert>",     lambda e: self._on_copy_files())
        self.tree.bind("<Button-3>",           self._on_right_click)

    # ── 로그 ─────────────────────────────────────────────────────
    def _log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.configure(state="disabled")
        self.log_text.see("end")

    # ── 툴팁 ──────────────────────────────────────────────────────
    def _on_hover(self, event):
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)

        if not row_id or row_id not in self._full_data:
            self._hide_tooltip()
            return

        full = self._full_data[row_id]
        col_map = {"#1": str(full[0]), "#2": full[1], "#3": full[2], "#4": full[3]}
        text = col_map.get(col_id, "")

        if not text or len(text) < 30:
            self._hide_tooltip()
            return

        if len(text) > 800:
            text = text[:800] + " …"

        # 미리보기 컬럼(#4) hover 시에만 키워드 강조. 폴더/파일명에는 의미 없음.
        keywords = _prepare_keywords(self._last_kw, self._last_mode) if col_id == "#4" else []
        self._show_tooltip(event, text, keywords)

    def _show_tooltip(self, event, text: str, keywords: list[str] = ()):
        self._hide_tooltip()
        tw = tk.Toplevel(self.root)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)

        # tk.Text 사용 이유: tag 로 키워드 부분만 빨간 굵게 강조 가능 (tk.Label 은 부분 색상 불가).
        width_chars  = 60
        height_lines = min(20, max(3, len(text) // width_chars + 2))
        txt = tk.Text(tw, width=width_chars, height=height_lines,
                      bg="#ffffe0", fg="#222",
                      relief=tk.SOLID, borderwidth=1,
                      font=("맑은 고딕", 9), wrap="word",
                      padx=6, pady=4, cursor="arrow",
                      highlightthickness=0)
        txt.insert("1.0", text)

        # 키워드가 있으면 텍스트 내 모든 출현을 빨간 굵게 (대소문자 무시).
        if keywords:
            txt.tag_configure("hl", foreground="red", font=("맑은 고딕", 9, "bold"))
            text_lower = text.lower()
            for kw in keywords:
                if not kw:
                    continue
                kw_lower, kw_len = kw.lower(), len(kw)
                start = 0
                while True:
                    pos = text_lower.find(kw_lower, start)
                    if pos < 0:
                        break
                    txt.tag_add("hl", f"1.0+{pos}c", f"1.0+{pos + kw_len}c")
                    start = pos + kw_len

        txt.configure(state="disabled")  # 읽기 전용 — 편집/삽입 방지
        txt.pack()

        x, y = event.x_root + 15, event.y_root + 10
        tw.update_idletasks()
        sw, sh     = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        tw_w, tw_h = tw.winfo_width(), tw.winfo_height()
        if x + tw_w > sw: x = sw - tw_w - 10
        if y + tw_h > sh: y = event.y_root - tw_h - 10
        tw.wm_geometry(f"+{x}+{y}")
        self._tooltip = tw

    def _hide_tooltip(self, event=None):
        if self._tooltip:
            self._tooltip.destroy()
            self._tooltip = None

    # ── 검색 ──────────────────────────────────────────────────────
    def _insert_rows(self, rows):
        # 마지막 검색의 키워드 리스트로 스니펫을 추출 — 키워드 인근 ±SNIPPET_RADIUS 자.
        kws = _prepare_keywords(self._last_kw, self._last_mode)
        for row in rows:
            rid, directory, filename, body_chunk = row
            is_excluded   = not body_chunk  # NULL 또는 빈 문자열
            preview_full  = "" if is_excluded else _extract_snippet(body_chunk or "", kws)
            preview_short = preview_full[:150] + "…" if len(preview_full) > 150 else preview_full
            dir_short     = directory
            fn_short      = filename[:50]  + "…" if len(filename) > 53  else filename

            tags = ("excluded",) if is_excluded else ()
            self.tree.insert("", tk.END, iid=str(rid),
                             values=(rid, dir_short, fn_short, preview_short), tags=tags)
            self._full_data[str(rid)] = (rid, directory, filename, preview_full)
            if is_excluded:
                self._excluded_ids.add(str(rid))
        self.results.extend(rows)

    def _update_status(self):
        shown     = len(self.results)
        remaining = self._total - shown
        if remaining > 0:
            self.status.configure(text=f"{self._total:,}건 중 {shown:,}건 표시 (잔여 {remaining:,})")
            self.more_btn.configure(state=tk.NORMAL)
            self.all_btn.configure(state=tk.NORMAL)
        else:
            self.status.configure(text=f"{self._total:,}건 (전체)")
            self.more_btn.configure(state=tk.DISABLED)
            self.all_btn.configure(state=tk.DISABLED)

    def _on_toggle_id_filter(self):
        state = tk.NORMAL if self.id_filter_var.get() else tk.DISABLED
        self.id_min_entry.configure(state=state)
        self.id_max_entry.configure(state=state)

    def _parse_id_filters(self):
        if not self.id_filter_var.get():
            return None, None
        def _p(s: str):
            s = s.strip()
            if not s:
                return None
            try:
                return int(s)
            except ValueError:
                raise ValueError(f"ID 값이 정수가 아닙니다: {s!r}")
        return _p(self.id_min_var.get()), _p(self.id_max_var.get())

    def _on_search(self):
        kw = self.entry.get().strip()
        self.btn.configure(state=tk.DISABLED)
        self.status.configure(text="검색 중...")
        self.root.update()
        try:
            target           = self.target_var.get()
            mode             = self.mode_var.get()
            include_excluded = self.include_excluded_var.get()
            id_min, id_max   = self._parse_id_filters()
            self._last_kw               = kw
            self._last_target           = target
            self._last_mode             = mode
            self._last_include_excluded = include_excluded
            self._last_id_min           = id_min
            self._last_id_max           = id_max
            self._offset      = 0
            self._total       = count_results(kw, target, mode, include_excluded,
                                              id_min=id_min, id_max=id_max)
            self.results      = []
            self.tree.delete(*self.tree.get_children())
            self._full_data   = {}
            self._excluded_ids = set()

            mode_label = {"and": "AND", "or": "OR", "phrase": "전체문자열"}[mode]
            excl_label = " [제외 포함]" if include_excluded else ""
            id_label = ""
            if id_min is not None or id_max is not None:
                id_label = f" [ID {id_min if id_min is not None else ''}~{id_max if id_max is not None else ''}]"
            self._log(f"[검색] '{kw}' | 대상: {target} | 방식: {mode_label}{excl_label}{id_label} → {self._total:,}건")

            rows = search(kw, target, mode, offset=0, include_excluded=include_excluded,
                          id_min=id_min, id_max=id_max)
            self._offset = len(rows)
            self._insert_rows(rows)
            self._update_status()
        except Exception as e:
            messagebox.showerror("오류", str(e))
            self.status.configure(text="오류")
            self._log(f"[오류] {e}")
        finally:
            self.btn.configure(state=tk.NORMAL)

    def _on_more(self):
        self.more_btn.configure(state=tk.DISABLED)
        self.status.configure(text="추가 로딩...")
        self.root.update()
        try:
            rows = search(self._last_kw, self._last_target, self._last_mode,
                          offset=self._offset, include_excluded=self._last_include_excluded,
                          id_min=self._last_id_min, id_max=self._last_id_max)
            self._offset += len(rows)
            self._insert_rows(rows)
            self._update_status()
        except Exception as e:
            messagebox.showerror("오류", str(e))

    def _on_load_all(self):
        self.all_btn.configure(state=tk.DISABLED)
        self.more_btn.configure(state=tk.DISABLED)
        remaining = self._total - len(self.results)
        self.status.configure(text=f"전체 로딩 중 ({remaining:,}건)...")
        self.root.update()
        try:
            rows = search(self._last_kw, self._last_target, self._last_mode,
                          limit=remaining, offset=self._offset,
                          include_excluded=self._last_include_excluded,
                          id_min=self._last_id_min, id_max=self._last_id_max)
            self._offset += len(rows)
            self._insert_rows(rows)
            self._update_status()
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # ── 키보드 확장 선택 ──────────────────────────────────────────
    def _reset_anchor_on_move(self, event=None):
        _ = event
        self._sel_anchor = None  # 다음 Shift+방향 시 현재 focus를 앵커로 사용

    def _extend_selection(self, direction: int):
        children = self.tree.get_children()
        if not children:
            return "break"

        cur = self.tree.focus()
        if not cur or not self.tree.exists(cur):
            cur = children[0] if direction > 0 else children[-1]
            self._sel_anchor = cur
            self.tree.selection_set(cur)
            self.tree.focus(cur)
            self.tree.see(cur)
            return "break"

        if self._sel_anchor is None or not self.tree.exists(self._sel_anchor):
            self._sel_anchor = cur

        nxt = self.tree.next(cur) if direction > 0 else self.tree.prev(cur)
        if not nxt:
            return "break"

        idx_a = children.index(self._sel_anchor) if self._sel_anchor in children else children.index(cur)
        idx_n = children.index(nxt)
        lo, hi = (idx_a, idx_n) if idx_a <= idx_n else (idx_n, idx_a)
        self.tree.selection_set(children[lo:hi + 1])
        self.tree.focus(nxt)
        self.tree.see(nxt)
        return "break"

    # ── 선택 / 열기 ──────────────────────────────────────────────
    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            self.open_btn.configure(state=tk.DISABLED)
            self.del_btn.configure(state=tk.DISABLED, text="제외 (Del)")
            return

        excluded_sel = [iid for iid in sel if iid in self._excluded_ids]
        normal_sel   = [iid for iid in sel if iid not in self._excluded_ids]
        mixed        = bool(excluded_sel) and bool(normal_sel)

        if mixed:
            self.del_btn.configure(state=tk.DISABLED, text="혼합 선택 불가")
        elif excluded_sel:
            self.del_btn.configure(state=tk.NORMAL,   text="완전 삭제 (Del)")
        else:
            self.del_btn.configure(state=tk.NORMAL,   text="제외 (Del)")

        if len(sel) == 1:
            rid = sel[0]
            if rid in self._full_data:
                _, directory, filename, _ = self._full_data[rid]
                fp      = os.path.join(directory, filename)
                display = fp if len(fp) <= 90 else fp[:87] + "…"
                self.info_label.configure(text=display, foreground="black")
            self.open_btn.configure(state=tk.NORMAL)
            self.open_path_btn.configure(state=tk.NORMAL)
        else:
            self.info_label.configure(text=f"{len(sel)}건 선택됨", foreground="black")
            self.open_btn.configure(state=tk.DISABLED)
            self.open_path_btn.configure(state=tk.DISABLED)

    def _on_open(self):
        sel = self.tree.selection()
        if not sel or len(sel) != 1:
            return
        rid = sel[0]
        if rid in self._full_data:
            _, directory, filename, _ = self._full_data[rid]
            fp = os.path.join(directory, filename)
            if not os.path.exists(fp):
                messagebox.showwarning("파일 없음", f"파일을 찾을 수 없습니다:\n{fp}")
                return
            try:
                os.startfile(fp)
            except Exception as e:
                messagebox.showerror("열기 실패", str(e))

    def _on_open_path(self):
        """탐색기를 띄워 해당 파일이 선택된 상태로 폴더를 연다.
        파일이 없으면 폴더만 연다(폴더도 없으면 경고)."""
        sel = self.tree.selection()
        if not sel or len(sel) != 1:
            return
        rid = sel[0]
        if rid not in self._full_data:
            return
        _, directory, filename, _ = self._full_data[rid]
        fp = os.path.join(directory, filename)
        try:
            if os.path.exists(fp):
                # explorer.exe /select,<path> 는 리스트 인자로 넘기면 CreateProcess 의
                # quoting 단계에서 따옴표가 끼어들어 explorer 의 특이한 파서가 경로를
                # 못 읽고 기본 폴더(문서) 로 폴백한다. 슬래시가 섞여도 같은 증상.
                # normpath 로 백슬래시 통일 + 단일 문자열로 전달해야 안정적.
                fp_norm = os.path.normpath(fp)
                subprocess.Popen(f'explorer /select,"{fp_norm}"')
            elif os.path.isdir(directory):
                os.startfile(os.path.normpath(directory))
            else:
                messagebox.showwarning("경로 없음", f"경로를 찾을 수 없습니다:\n{directory}")
        except Exception as e:
            messagebox.showerror("열기 실패", str(e))

    # ── 우클릭 컨텍스트 메뉴 ─────────────────────────────────────
    # _full_data 튜플 = (rid, directory, filename, preview_full).
    # 컬럼 ID(#1~#4) → (라벨, 튜플 인덱스). ID 컬럼(#1) 은 hover 도 안 뜨므로 제외.
    _COL_COPY_MAP = {
        "#2": ("폴더 경로 복사",     1),
        "#3": ("파일명 복사",         2),
        "#4": ("내용 미리보기 복사",  3),
    }

    def _on_right_click(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return  # 빈 영역 우클릭은 무시
        # 클릭한 행이 현재 선택에 없으면 그 행만 선택. 이미 다중 선택 중이면 유지.
        if row_id not in self.tree.selection():
            self.tree.selection_set(row_id)
            self.tree.focus(row_id)

        sel = self.tree.selection()
        is_single = len(sel) == 1
        state_single = tk.NORMAL if is_single else tk.DISABLED

        m = self.context_menu
        m.delete(0, "end")

        # 컬럼별 셀 복사 — hover 툴팁이 뜨는 컬럼과 한 쌍.
        col_id = self.tree.identify_column(event.x)
        col_copy = self._COL_COPY_MAP.get(col_id)
        if col_copy:
            label, field_idx = col_copy
            m.add_command(label=label, command=lambda i=field_idx: self._copy_field(i))
            m.add_separator()

        m.add_command(label="파일 열기", command=self._on_open,      state=state_single)
        m.add_command(label="경로 열기", command=self._on_open_path, state=state_single)
        m.add_separator()

        # 삭제 항목 — 하단 del_btn 과 동일 규칙 (_on_select 참조).
        excluded_sel = [iid for iid in sel if iid in self._excluded_ids]
        normal_sel   = [iid for iid in sel if iid not in self._excluded_ids]
        if excluded_sel and normal_sel:
            m.add_command(label="혼합 선택 불가",  state=tk.DISABLED)
        elif excluded_sel:
            m.add_command(label="완전 삭제 (Del)", command=self._on_delete)
        else:
            m.add_command(label="제외 (Del)",      command=self._on_delete)

        try:
            m.tk_popup(event.x_root, event.y_root)
        finally:
            m.grab_release()

    def _copy_field(self, field_idx: int):
        """선택된 행들의 _full_data[field_idx] 를 줄바꿈으로 합쳐 클립보드에 복사."""
        sel = self.tree.selection()
        if not sel:
            return
        values = []
        for iid in sel:
            full = self._full_data.get(iid)
            if not full:
                continue
            v = full[field_idx]
            if v is None:
                continue
            values.append(str(v))
        if not values:
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append("\n".join(values))
            self.root.update()
        except Exception as e:
            messagebox.showerror("복사 실패", str(e))

    # ── Ctrl+C — 클립보드로 파일 복사 (drag 우회용) ──────────────
    def _on_copy_files(self):
        """선택된 파일들을 Explorer 호환 CF_HDROP 포맷으로 클립보드에 올림.
        Explorer 의 폴더에서 Ctrl+V 하면 파일이 복사된다."""
        sel = self.tree.selection()
        if not sel:
            return "break"

        paths, missing = [], []
        for iid in sel:
            full = self._full_data.get(iid)
            if not full:
                continue
            _, directory, filename, _ = full
            fp = os.path.join(directory, filename)
            (paths if os.path.exists(fp) else missing).append(fp)

        if missing:
            self._log(f"[복사] 누락된 파일 {len(missing)}건 제외")
        if not paths:
            self._log("[복사] 복사 가능한 파일 없음")
            return "break"

        try:
            _copy_files_to_clipboard(paths)
            self._log(f"[복사] {len(paths)}건 클립보드에 올림 — Explorer 에서 Ctrl+V")
        except ImportError:
            self._log("[복사 실패] pywin32 미설치")
        except Exception as e:
            self._log(f"[복사 실패] {e}")
        return "break"

    # ── 드래그-앤-드롭 (탐색기로 파일 복사) ──────────────────────
    def _on_drag_init(self, event):
        _ = event  # DnDEvent — x/y 없음. 클릭 시점에 selection이 이미 갱신됨
        self._hide_tooltip()

        sel = list(self.tree.selection())
        if not sel:
            # 드물게 selection이 비어있다면 커서 위치(루트 좌표) 기준으로 식별
            try:
                ry = self.tree.winfo_pointery() - self.tree.winfo_rooty()
                row = self.tree.identify_row(ry)
                if row:
                    sel = [row]
                    self.tree.selection_set(row)
            except Exception:
                pass

        paths, missing = [], []
        for iid in sel:
            full = self._full_data.get(iid)
            if not full:
                continue
            _, directory, filename, _ = full
            fp = os.path.join(directory, filename)
            (paths if os.path.exists(fp) else missing).append(fp)

        if missing:
            self._log(f"[드래그] 누락된 파일 {len(missing)}건 제외")
        if not paths:
            return "break"

        self._log(f"[드래그] {len(paths)}건을 탐색기로 전달")
        return ("copy", DND_FILES, tuple(paths))

    # ── 제외 / 완전 삭제 ─────────────────────────────────────────
    def _on_delete(self):
        sel = self.tree.selection()
        if not sel:
            return

        excluded_sel = [iid for iid in sel if iid in self._excluded_ids]
        normal_sel   = [iid for iid in sel if iid not in self._excluded_ids]
        if excluded_sel and normal_sel:
            return  # 혼합 선택 — 버튼이 이미 비활성이라 여기 올 일 없음

        is_hard_delete = bool(excluded_sel)
        target_iids    = excluded_sel if is_hard_delete else list(sel)

        ids, names = [], []
        for iid in target_iids:
            if iid in self._full_data:
                rid, _, filename, _ = self._full_data[iid]
                ids.append(rid)
                names.append(filename)

        if not ids:
            return

        preview = "\n".join(names[:10])
        if len(names) > 10:
            preview += f"\n… 외 {len(names) - 10}건"

        if is_hard_delete:
            confirmed = messagebox.askyesno(
                "완전 삭제 확인",
                f"{len(ids)}건을 DB에서 완전히 삭제합니다.\n(레코드 자체가 제거되며, 로컬 파일은 유지됩니다)\n\n{preview}\n\n계속하시겠습니까?",
            )
            action_label = "완전 삭제"
        else:
            confirmed = messagebox.askyesno(
                "제외 확인",
                f"{len(ids)}건을 검색에서 제외합니다.\n(레코드는 유지되며 body_text만 초기화됩니다)\n\n{preview}\n\n계속하시겠습니까?",
            )
            action_label = "제외"

        if not confirmed:
            self._log(f"[{action_label} 취소]")
            return

        id_preview = str(ids[:10]) + ("..." if len(ids) > 10 else "")
        self._log(f"[{action_label} 요청] {len(ids)}건 | ID: {id_preview}")
        self.del_btn.configure(state=tk.DISABLED)

        threading.Thread(
            target=self._do_delete,
            args=(ids, tuple(target_iids), is_hard_delete),
            daemon=True,
        ).start()

    def _do_delete(self, ids: list, iids: tuple, is_hard_delete: bool):
        try:
            affected = delete_rows(ids) if is_hard_delete else nullify_body_text(ids)
            self.root.after(0, lambda: self._after_delete(iids, affected, is_hard_delete))
        except Exception as e:
            action = "완전 삭제" if is_hard_delete else "제외"
            self.root.after(0, lambda err=e: self._log(f"[오류] {action} 처리 실패: {err}"))
            self.root.after(0, lambda: self.del_btn.configure(state=tk.NORMAL))

    def _after_delete(self, iids: tuple, affected: int, is_hard_delete: bool):
        iid_set = set(iids)

        next_focus = ""
        last_iid = iids[-1] if iids else ""
        if last_iid and self.tree.exists(last_iid):
            cand = self.tree.next(last_iid)
            while cand and cand in iid_set:
                cand = self.tree.next(cand)
            if not cand:
                first_iid = iids[0]
                cand = self.tree.prev(first_iid) if self.tree.exists(first_iid) else ""
                while cand and cand in iid_set:
                    cand = self.tree.prev(cand)
            next_focus = cand or ""

        for iid in iids:
            self._full_data.pop(iid, None)
            self._excluded_ids.discard(iid)
            if self.tree.exists(iid):
                self.tree.delete(iid)

        if next_focus and self.tree.exists(next_focus):
            self.tree.selection_set(next_focus)
            self.tree.focus(next_focus)
            self.tree.see(next_focus)
            self.tree.focus_set()

        self._total   = max(0, self._total - affected)
        self.results  = [r for r in self.results if str(r[0]) not in iid_set]
        self._offset  = max(0, self._offset - affected)

        self._update_status()
        if is_hard_delete:
            self._log(f"[완전 삭제 완료] {affected}건 DB에서 삭제됨 (로컬 파일 유지)")
        else:
            self._log(f"[제외 완료] {affected}건 검색에서 제외됨 (레코드 유지)")
        self.del_btn.configure(state=tk.DISABLED, text="제외 (Del)")
        self.info_label.configure(text=self._info_default, foreground="gray")


def main():
    if _DND_AVAILABLE:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            root = tk.Tk()
    else:
        root = tk.Tk()
    style = ttk.Style()
    style.configure("Treeview",         font=("맑은 고딕", 9), rowheight=24)
    style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"))
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
