"""
Doc Mine — 통합 GUI 런처
==========================
4단계(스캔 / 적재 / 검색 / 추출) 를 단일 창의 탭으로 통합.

- 콘솔 출력(스캔 / 적재) 은 sys.stdout 을 리다이렉트해 각 탭의 로그 위젯에
  실시간 표시. inserter.PB 의 \\r 진행바도 단일 라이브 라인으로 갱신.
- 검색 / 추출 탭은 기존 search_gui.App / extractor_gui.ExtractorApp 을
  Notebook 탭의 ttk.Frame 에 임베드.

단독 실행:
    python -m docmine.unified_gui
또는
    docmine g
"""
from __future__ import annotations

import contextlib
import multiprocessing as mp
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from . import __app_name__, config
from .about import show_about
from .drive_picker import list_drives
from .icon import make_app_icon

try:
    from tkinterdnd2 import TkinterDnD
    _DND_AVAILABLE = True
except ImportError:
    TkinterDnD = None  # type: ignore[assignment]
    _DND_AVAILABLE = False


# ═══════════════════════════════════════════════════════════════
# stdout 리다이렉트 (스캔 / 적재 탭용)
# ═══════════════════════════════════════════════════════════════

class _QueueWriter:
    """sys.stdout/sys.stderr 대체 — 라인/CR/잔여 partial 단위로 큐에 push."""

    def __init__(self, q: "queue.Queue", mirror=None):
        self.q = q
        self.mirror = mirror
        self._buf = ""

    def write(self, s: str) -> int:
        if not s:
            return 0
        if self.mirror is not None:
            try:
                self.mirror.write(s)
            except Exception:
                pass
        self._buf += s
        while True:
            i_nl = self._buf.find("\n")
            i_cr = self._buf.find("\r")
            cuts = [i for i in (i_nl, i_cr) if i >= 0]
            if not cuts:
                break
            cut = min(cuts)
            line = self._buf[:cut]
            sep = self._buf[cut]
            self._buf = self._buf[cut + 1:]
            self.q.put((sep, line))
        if self._buf:
            # 줄이 아직 안 끝났어도 라이브 라인으로 즉시 반영 (\r 진행바용).
            self.q.put(("partial", self._buf))
        return len(s)

    def flush(self) -> None:
        if self.mirror is not None:
            try:
                self.mirror.flush()
            except Exception:
                pass

    def isatty(self) -> bool:
        return False


class _LogPane:
    """ScrolledText + 큐 드레인 — \\r/partial 은 'live line' 마크에서 덮어쓰기."""

    POLL_MS = 50

    def __init__(self, parent: tk.Misc):
        self.frame = ttk.LabelFrame(parent, text="로그", padding=4)
        self.text = scrolledtext.ScrolledText(
            self.frame, state="disabled",
            font=("Consolas", 9), wrap="word",
            bg="#1e1e1e", fg="#d4d4d4",
        )
        self.text.pack(fill=tk.BOTH, expand=True)
        self.text.configure(state=tk.NORMAL)
        self.text.mark_set("live", "end-1c")
        self.text.mark_gravity("live", "left")
        self.text.configure(state=tk.DISABLED)

        self.queue: "queue.Queue" = queue.Queue()
        self._poll_id = None

    def start_polling(self):
        if self._poll_id is None:
            self._poll_id = self.frame.after(self.POLL_MS, self._drain)

    def stop_polling(self):
        if self._poll_id is not None:
            try:
                self.frame.after_cancel(self._poll_id)
            except Exception:
                pass
            self._poll_id = None

    def clear(self):
        self.text.configure(state=tk.NORMAL)
        self.text.delete("1.0", "end")
        self.text.mark_set("live", "end-1c")
        self.text.configure(state=tk.DISABLED)

    def _drain(self):
        try:
            while True:
                sep, line = self.queue.get_nowait()
                self._append(sep, line)
        except queue.Empty:
            pass
        finally:
            self._poll_id = self.frame.after(self.POLL_MS, self._drain)

    def _append(self, sep: str, line: str):
        self.text.configure(state=tk.NORMAL)
        if sep == "\n":
            # 라이브 라인 갱신 후 newline — 마크를 다음 줄로 진전.
            self.text.delete("live", "end-1c")
            self.text.insert("live", line + "\n")
            self.text.mark_set("live", "end-1c")
        else:  # "\r" 또는 "partial" — 마크 위치 유지, 같은 줄 덮어쓰기.
            self.text.delete("live", "end-1c")
            self.text.insert("live", line)
        self.text.see("end")
        self.text.configure(state=tk.DISABLED)


@contextlib.contextmanager
def _redirect_stdio(writer: _QueueWriter):
    """with 블록 동안 sys.stdout/stderr 을 writer 로 교체."""
    orig_out, orig_err = sys.stdout, sys.stderr
    sys.stdout = writer
    sys.stderr = writer
    try:
        yield
    finally:
        sys.stdout = orig_out
        sys.stderr = orig_err


# ═══════════════════════════════════════════════════════════════
# Tab 1 — 스캔
# ═══════════════════════════════════════════════════════════════

class ScanTab:
    """HWP 또는 PDF 스캔 탭 — `scope` 파라미터로 분기.

    scope='hwp' : 확장자 .hwp/.hwpx, 출력 CSV = config.CSV_FILE
    scope='pdf' : 확장자 .pdf,        출력 CSV = config.PDF_CSV_FILE

    스캐너 코어(`scanner.run`) 자체는 공용이지만, 결과 CSV 와 다운스트림 적재
    테이블이 HWP/PDF 로 완전히 분리돼 있어 탭도 별도로 노출한다.
    """

    SCOPE_PRESETS = {
        "hwp": {
            "label":       "HWP",
            "ext_choices": [(".hwp", True), (".hwpx", True)],
            "default_csv": config.CSV_FILE,
        },
        "pdf": {
            "label":       "PDF",
            "ext_choices": [(".pdf", True)],
            "default_csv": config.PDF_CSV_FILE,
        },
    }

    def __init__(self, parent: tk.Misc, scope: str = "hwp"):
        if scope not in self.SCOPE_PRESETS:
            raise ValueError(f"unknown scope: {scope!r}")
        self.scope = scope
        preset = self.SCOPE_PRESETS[scope]
        self.scope_label = preset["label"]
        self._ext_choices = preset["ext_choices"]
        self._default_csv = preset["default_csv"]

        self.frame = ttk.Frame(parent, padding=8)
        self._busy = False
        self._build()

    def _build(self):
        # ── 확장자 선택 ──────────────────────────────────────────
        ext_lf = ttk.LabelFrame(
            self.frame, text=f"{self.scope_label} 대상 확장자", padding=8,
        )
        ext_lf.pack(fill=tk.X)
        self._ext_vars: list[tuple[tk.BooleanVar, str]] = []
        for ext, default_on in self._ext_choices:
            v = tk.BooleanVar(value=default_on)
            ttk.Checkbutton(ext_lf, text=ext, variable=v).pack(side=tk.LEFT, padx=(0, 16))
            self._ext_vars.append((v, ext))

        # ── 드라이브 선택 ────────────────────────────────────────
        drv_lf = ttk.LabelFrame(self.frame, text="스캔 대상 드라이브", padding=8)
        drv_lf.pack(fill=tk.X, pady=(8, 0))

        self._drive_vars: list[tuple[tk.BooleanVar, str]] = []
        default_set = {d.upper().rstrip("\\") + "\\" for d in config.SCAN_DRIVES}
        try:
            drives = list_drives()
        except Exception as e:
            drives = []
            ttk.Label(drv_lf, text=f"(드라이브 열거 실패: {e})",
                      foreground="red").pack(anchor="w")

        if not drives:
            ttk.Label(drv_lf, text="(마운트된 드라이브가 없습니다)",
                      foreground="gray").pack(anchor="w")
        for path, label, dtype in drives:
            v = tk.BooleanVar(value=(path.upper() in default_set))
            text = f"{path}    {label or '(라벨 없음)'}    [{dtype}]"
            ttk.Checkbutton(drv_lf, text=text, variable=v).pack(anchor="w")
            self._drive_vars.append((v, path))

        quick = ttk.Frame(drv_lf)
        quick.pack(anchor="w", pady=(6, 0))
        ttk.Button(quick, text="전체 선택",
                   command=lambda: self._set_all(True)).pack(side=tk.LEFT)
        ttk.Button(quick, text="전체 해제",
                   command=lambda: self._set_all(False)).pack(side=tk.LEFT, padx=(6, 0))

        # ── 출력 경로 ────────────────────────────────────────────
        out_lf = ttk.LabelFrame(
            self.frame, text=f"{self.scope_label} 출력 CSV", padding=8,
        )
        out_lf.pack(fill=tk.X, pady=(8, 0))
        self.out_var = tk.StringVar(value=str(Path(self._default_csv).absolute()))
        ttk.Entry(out_lf, textvariable=self.out_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(out_lf, text="저장 위치…",
                   command=self._browse_out).pack(side=tk.LEFT, padx=(6, 0))

        # ── 실행 버튼 ────────────────────────────────────────────
        btns = ttk.Frame(self.frame)
        btns.pack(fill=tk.X, pady=(8, 0))
        self.start_btn = ttk.Button(
            btns, text=f"{self.scope_label} 스캔 시작", command=self._on_start,
        )
        self.start_btn.pack(side=tk.LEFT)

        # ── 로그 ────────────────────────────────────────────────
        self.log = _LogPane(self.frame)
        self.log.frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.log.start_polling()

    def _set_all(self, value: bool):
        for v, _ in self._drive_vars:
            v.set(value)

    def _browse_out(self):
        path = filedialog.asksaveasfilename(
            title=f"{self.scope_label} 스캔 결과 CSV 저장 위치",
            defaultextension=".csv",
            initialfile=Path(self.out_var.get()).name or Path(self._default_csv).name,
            filetypes=[("CSV", "*.csv"), ("모든 파일", "*.*")],
        )
        if path:
            self.out_var.set(path)

    def _on_start(self):
        if self._busy:
            return
        drives = [p for v, p in self._drive_vars if v.get()]
        if not drives:
            messagebox.showwarning("드라이브 선택", "스캔할 드라이브를 하나 이상 선택하세요.")
            return
        exts = {ext for v, ext in self._ext_vars if v.get()}
        if not exts:
            messagebox.showwarning("확장자 선택", "대상 확장자를 하나 이상 선택하세요.")
            return
        out = self.out_var.get().strip()
        if not out:
            messagebox.showwarning("출력 경로", "출력 CSV 경로를 지정하세요.")
            return

        self._busy = True
        self.start_btn.configure(
            state=tk.DISABLED, text=f"{self.scope_label} 스캔 중…",
        )
        self.log.clear()

        writer = _QueueWriter(self.log.queue, mirror=sys.__stdout__)

        def _runner():
            try:
                from . import scanner
                with _redirect_stdio(writer):
                    scanner.run(drives=drives, out=out, extensions=exts)
            except Exception as e:
                writer.write(f"\n[오류] {type(e).__name__}: {e}\n")
            finally:
                writer.flush()
                self.frame.after(0, self._on_done)

        threading.Thread(target=_runner, daemon=True).start()

    def _on_done(self):
        self._busy = False
        self.start_btn.configure(
            state=tk.NORMAL, text=f"{self.scope_label} 스캔 시작",
        )


# ═══════════════════════════════════════════════════════════════
# Tab 2 — HWP 적재
# ═══════════════════════════════════════════════════════════════

class InsertTab:
    def __init__(self, parent: tk.Misc):
        self.frame = ttk.Frame(parent, padding=8)
        self._busy = False
        self._build()

    def _build(self):
        info = ttk.Label(
            self.frame,
            text="HWP/HWPX 본문을 한/글 COM 으로 파싱해 HWP 전용 테이블에 적재합니다.\n"
                 "(.pdf 행이 섞여 있어도 무시되며, PDF 는 별도 탭에서 적재하세요.)",
            foreground="gray",
            justify=tk.LEFT,
        )
        info.pack(fill=tk.X, pady=(0, 4))

        # ── CSV ─────────────────────────────────────────────────
        csv_lf = ttk.LabelFrame(self.frame, text="입력 CSV (HWP 스캔 결과)", padding=8)
        csv_lf.pack(fill=tk.X)
        self.csv_var = tk.StringVar(value=str(Path(config.CSV_FILE).absolute()))
        ttk.Entry(csv_lf, textvariable=self.csv_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(csv_lf, text="찾아보기…",
                   command=self._browse_csv).pack(side=tk.LEFT, padx=(6, 0))

        # ── 범위 ────────────────────────────────────────────────
        rng_lf = ttk.LabelFrame(self.frame, text="처리 범위 (비우면 전체)", padding=8)
        rng_lf.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(rng_lf, text="start").pack(side=tk.LEFT)
        self.start_var = tk.StringVar(value="0")
        ttk.Entry(rng_lf, textvariable=self.start_var, width=10).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(rng_lf, text="end").pack(side=tk.LEFT)
        self.end_var = tk.StringVar(value="")
        ttk.Entry(rng_lf, textvariable=self.end_var, width=10).pack(side=tk.LEFT, padx=(4, 0))

        # ── 실행 버튼 ────────────────────────────────────────────
        btns = ttk.Frame(self.frame)
        btns.pack(fill=tk.X, pady=(8, 0))
        self.start_btn = ttk.Button(btns, text="HWP 적재 시작", command=self._on_start)
        self.start_btn.pack(side=tk.LEFT)

        # ── 로그 ────────────────────────────────────────────────
        self.log = _LogPane(self.frame)
        self.log.frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.log.start_polling()

    def _browse_csv(self):
        path = filedialog.askopenfilename(
            title="HWP CSV 선택",
            filetypes=[("CSV", "*.csv"), ("모든 파일", "*.*")],
            initialfile=Path(self.csv_var.get()).name or Path(config.CSV_FILE).name,
        )
        if path:
            self.csv_var.set(path)

    def _parse_int(self, s: str, default):
        s = s.strip()
        if not s:
            return default
        try:
            return int(s)
        except ValueError:
            raise ValueError(f"정수가 아닙니다: {s!r}")

    def _on_start(self):
        if self._busy:
            return
        csv_path = self.csv_var.get().strip()
        if not csv_path or not Path(csv_path).exists():
            messagebox.showwarning("CSV 없음",
                                   f"CSV 파일을 찾을 수 없습니다:\n{csv_path}")
            return
        try:
            start = self._parse_int(self.start_var.get(), 0)
            end = self._parse_int(self.end_var.get(), None)
        except ValueError as e:
            messagebox.showwarning("범위 오류", str(e))
            return

        self._busy = True
        self.start_btn.configure(state=tk.DISABLED, text="HWP 적재 중…")
        self.log.clear()

        writer = _QueueWriter(self.log.queue, mirror=sys.__stdout__)

        def _runner():
            try:
                from . import inserter
                from .hwp_parser import configure_logging
                configure_logging(verbose=False)
                with _redirect_stdio(writer):
                    inserter.run(csv_path, start=start, end=end)
            except Exception as e:
                writer.write(f"\n[오류] {type(e).__name__}: {e}\n")
            finally:
                writer.flush()
                self.frame.after(0, self._on_done)

        threading.Thread(target=_runner, daemon=True).start()

    def _on_done(self):
        self._busy = False
        self.start_btn.configure(state=tk.NORMAL, text="HWP 적재 시작")


# ═══════════════════════════════════════════════════════════════
# Tab 2' — PDF 적재
# ═══════════════════════════════════════════════════════════════

class PdfInsertTab:
    """InsertTab 과 동일 구조 — pdf_inserter.run 을 호출.

    HWP 와 완전히 분리된 PDF 전용 테이블(`config.PDF_DB_TABLE`) 로 적재한다.
    PyMuPDF 기반이라 스캔본/이미지 PDF 는 빈 본문이 나올 수 있으며(OCR 미적용),
    그런 경우 'error' 상태로 기록된다.
    """

    def __init__(self, parent: tk.Misc):
        self.frame = ttk.Frame(parent, padding=8)
        self._busy = False
        self._build()

    def _build(self):
        info = ttk.Label(
            self.frame,
            text="PyMuPDF 로 PDF 본문 텍스트를 뽑아 PDF 전용 테이블에 적재합니다.\n"
                 "스캔본/이미지 PDF 는 빈 본문이 나올 수 있습니다 (OCR 미적용).",
            foreground="gray",
            justify=tk.LEFT,
        )
        info.pack(fill=tk.X, pady=(0, 4))

        # ── CSV ─────────────────────────────────────────────────
        csv_lf = ttk.LabelFrame(self.frame, text="입력 CSV (PDF 스캔 결과)", padding=8)
        csv_lf.pack(fill=tk.X)
        self.csv_var = tk.StringVar(value=str(Path(config.PDF_CSV_FILE).absolute()))
        ttk.Entry(csv_lf, textvariable=self.csv_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(csv_lf, text="찾아보기…",
                   command=self._browse_csv).pack(side=tk.LEFT, padx=(6, 0))

        # ── 범위 ────────────────────────────────────────────────
        rng_lf = ttk.LabelFrame(self.frame, text="처리 범위 (비우면 전체)", padding=8)
        rng_lf.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(rng_lf, text="start").pack(side=tk.LEFT)
        self.start_var = tk.StringVar(value="0")
        ttk.Entry(rng_lf, textvariable=self.start_var, width=10).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(rng_lf, text="end").pack(side=tk.LEFT)
        self.end_var = tk.StringVar(value="")
        ttk.Entry(rng_lf, textvariable=self.end_var, width=10).pack(side=tk.LEFT, padx=(4, 0))

        # ── 실행 버튼 ────────────────────────────────────────────
        btns = ttk.Frame(self.frame)
        btns.pack(fill=tk.X, pady=(8, 0))
        self.start_btn = ttk.Button(btns, text="PDF 적재 시작", command=self._on_start)
        self.start_btn.pack(side=tk.LEFT)

        # ── 로그 ────────────────────────────────────────────────
        self.log = _LogPane(self.frame)
        self.log.frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.log.start_polling()

    def _browse_csv(self):
        path = filedialog.askopenfilename(
            title="PDF CSV 선택",
            filetypes=[("CSV", "*.csv"), ("모든 파일", "*.*")],
            initialfile=Path(self.csv_var.get()).name or Path(config.PDF_CSV_FILE).name,
        )
        if path:
            self.csv_var.set(path)

    def _parse_int(self, s: str, default):
        s = s.strip()
        if not s:
            return default
        try:
            return int(s)
        except ValueError:
            raise ValueError(f"정수가 아닙니다: {s!r}")

    def _on_start(self):
        if self._busy:
            return
        csv_path = self.csv_var.get().strip()
        if not csv_path or not Path(csv_path).exists():
            messagebox.showwarning("CSV 없음",
                                   f"CSV 파일을 찾을 수 없습니다:\n{csv_path}")
            return
        try:
            start = self._parse_int(self.start_var.get(), 0)
            end = self._parse_int(self.end_var.get(), None)
        except ValueError as e:
            messagebox.showwarning("범위 오류", str(e))
            return

        self._busy = True
        self.start_btn.configure(state=tk.DISABLED, text="PDF 적재 중…")
        self.log.clear()

        writer = _QueueWriter(self.log.queue, mirror=sys.__stdout__)

        def _runner():
            try:
                from . import pdf_inserter
                with _redirect_stdio(writer):
                    pdf_inserter.run(csv_path, start=start, end=end)
            except Exception as e:
                writer.write(f"\n[오류] {type(e).__name__}: {e}\n")
            finally:
                writer.flush()
                self.frame.after(0, self._on_done)

        threading.Thread(target=_runner, daemon=True).start()

    def _on_done(self):
        self._busy = False
        self.start_btn.configure(state=tk.NORMAL, text="PDF 적재 시작")


# ═══════════════════════════════════════════════════════════════
# 메인 윈도우
# ═══════════════════════════════════════════════════════════════

class UnifiedApp:
    def __init__(self, root: tk.Misc):
        self.root = root
        root.title(__app_name__)
        root.geometry("1180x820")
        root.minsize(900, 600)

        # 아이콘 (GC 방지용 인스턴스 속성)
        try:
            self._app_icon = make_app_icon(root)
            root.iconphoto(True, self._app_icon)
        except tk.TclError:
            self._app_icon = None

        style = ttk.Style()
        style.configure("Treeview", font=("맑은 고딕", 9), rowheight=24)
        style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"))
        style.configure("TLabelframe.Label", font=("맑은 고딕", 9, "bold"))

        # ── 상단 바: About 버튼만 (앱 이름/버전은 윈도우 제목 + ? 다이얼로그에서) ──
        topbar = ttk.Frame(root, padding=(10, 6, 10, 0))
        topbar.pack(fill=tk.X)
        ttk.Button(topbar, text="?", width=3,
                   command=lambda: show_about(root)).pack(side=tk.RIGHT)

        # ── Notebook ────────────────────────────────────────────
        # HWP / PDF 파이프라인을 같은 단계끼리 나란히 노출 — 사용자가 어느
        # 포맷을 작업 중인지 항상 명시적으로 선택하도록.
        nb = ttk.Notebook(root)
        nb.pack(fill=tk.BOTH, expand=True, padx=10, pady=(6, 10))

        self.scan_tab     = ScanTab(nb, scope="hwp")
        nb.add(self.scan_tab.frame,     text="① HWP 스캔")

        self.pdf_scan_tab = ScanTab(nb, scope="pdf")
        nb.add(self.pdf_scan_tab.frame, text="① PDF 스캔")

        self.insert_tab     = InsertTab(nb)
        nb.add(self.insert_tab.frame,     text="② HWP 적재")

        self.pdf_insert_tab = PdfInsertTab(nb)
        nb.add(self.pdf_insert_tab.frame, text="② PDF 적재")

        # 검색 — HWP/PDF 적재 결과를 한 테이블에서 통합 검색
        from . import search_gui
        search_frame = ttk.Frame(nb)
        nb.add(search_frame, text="③ 검색 (HWP+PDF)")
        self.search_app = search_gui.App(search_frame)

        # 추출 — HWP 전용 (한/글 COM 워커 사용)
        from . import extractor_gui
        extract_frame = ttk.Frame(nb)
        nb.add(extract_frame, text="④ HWP 추출")
        self.extract_app = extractor_gui.ExtractorApp(extract_frame)

        root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        busy_tabs = (
            self.scan_tab, self.pdf_scan_tab,
            self.insert_tab, self.pdf_insert_tab,
        )
        if any(t._busy for t in busy_tabs):
            if not messagebox.askyesno(
                "진행 중 작업",
                "스캔/적재 작업이 진행 중입니다.\n그래도 창을 닫겠습니까?",
            ):
                return
        self.root.destroy()


# ═══════════════════════════════════════════════════════════════
# 진입점
# ═══════════════════════════════════════════════════════════════

def main():
    mp.freeze_support()
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except Exception:
            pass

    # DnD 지원: search 탭의 드래그가 동작하려면 root 가 TkinterDnD.Tk 여야 함.
    if _DND_AVAILABLE:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            root = tk.Tk()
    else:
        root = tk.Tk()

    UnifiedApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
