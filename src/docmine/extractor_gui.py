"""
HWP / HWPX / PDF → TXT 변환기
================================
파일 1개 또는 폴더 내 HWP/HWPX/PDF 를 파싱하여
하나의 TXT 파일로 추출합니다. HWP/PDF 는 체크박스로 개별 선택.

단독 실행:
  python extractor_gui.py
"""

import os
import sys
import time
import multiprocessing as mp
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from queue import Empty

from .inserter import worker_main
from . import config
from .about import show_about
from .icon import make_app_icon

HWP_EXTS = {".hwp", ".hwpx"}
PDF_EXTS = {".pdf"}
SEPARATOR  = "=" * 80


# ═══════════════════════════════════════════════════════════════
# 파일 수집
# ═══════════════════════════════════════════════════════════════

def collect_files(path: str, single: bool, exts: set[str]) -> list[Path]:
    p = Path(path)
    if single:
        return [p] if p.suffix.lower() in exts else []
    return sorted(
        f for f in p.rglob("*") if f.suffix.lower() in exts
    )


# ═══════════════════════════════════════════════════════════════
# 워커
# ═══════════════════════════════════════════════════════════════

def _spawn_worker(task_q, result_q):
    # kill_hwp=False — 외부에서 띄워둔 한/글이 같이 종료되지 않도록
    # 워커 내부의 _kill_hwp() 호출을 막는다.
    w = mp.Process(target=worker_main, args=(task_q, result_q, False), daemon=True)
    w.start()
    return w


# ═══════════════════════════════════════════════════════════════
# GUI
# ═══════════════════════════════════════════════════════════════

class ExtractorApp:
    def __init__(self, master: tk.Misc):
        # master 가 Tk/Toplevel 이면 단독 창 모드, 그 외(Notebook 탭의 Frame 등)는 임베드 모드.
        self.root = master
        self._standalone = isinstance(master, (tk.Tk, tk.Toplevel))

        if self._standalone:
            master.title("HWP / HWPX → TXT 변환기")
            master.geometry("720x560")
            master.minsize(600, 480)
            master.resizable(True, True)

            # 아이콘은 PhotoImage GC 방지를 위해 인스턴스 속성으로 보관
            try:
                self._app_icon = make_app_icon(master)
                master.iconphoto(True, self._app_icon)
            except tk.TclError:
                self._app_icon = None

        self._running   = False
        self._stop_flag = False

        self._build_ui()

    # ── UI 구성 ───────────────────────────────────────────────

    def _build_ui(self):
        pad = {"padx": 12, "pady": 6}

        # ── 모드 ─────────────────────────────────────────────
        mode_frame = ttk.LabelFrame(self.root, text="변환 대상", padding=8)
        mode_frame.pack(fill=tk.X, **pad)

        self.mode_var = tk.StringVar(value="folder")
        ttk.Radiobutton(
            mode_frame, text="폴더 전체 (하위 포함)",
            variable=self.mode_var, value="folder",
            command=self._on_mode_change,
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            mode_frame, text="파일 1개",
            variable=self.mode_var, value="file",
            command=self._on_mode_change,
        ).pack(side=tk.LEFT, padx=(20, 0))

        ttk.Button(mode_frame, text="?", width=3,
                   command=lambda: show_about(self.root)).pack(side=tk.RIGHT)

        # ── 포맷 선택 ────────────────────────────────────────
        fmt_frame = ttk.LabelFrame(self.root, text="포맷 선택", padding=8)
        fmt_frame.pack(fill=tk.X, **pad)

        self.hwp_var = tk.BooleanVar(value=True)
        self.pdf_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            fmt_frame, text="HWP / HWPX (한/글 COM 파싱)",
            variable=self.hwp_var,
        ).pack(side=tk.LEFT)
        ttk.Checkbutton(
            fmt_frame, text="PDF (PyMuPDF 텍스트 추출)",
            variable=self.pdf_var,
        ).pack(side=tk.LEFT, padx=(20, 0))

        # ── 입력 경로 ─────────────────────────────────────────
        src_frame = ttk.LabelFrame(self.root, text="입력 경로", padding=8)
        src_frame.pack(fill=tk.X, **pad)

        self.src_var = tk.StringVar()
        self.src_entry = ttk.Entry(src_frame, textvariable=self.src_var, font=("맑은 고딕", 9))
        self.src_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.src_btn = ttk.Button(src_frame, text="찾아보기…", command=self._browse_src)
        self.src_btn.pack(side=tk.LEFT, padx=(6, 0))

        # ── 출력 파일 ─────────────────────────────────────────
        dst_frame = ttk.LabelFrame(self.root, text="저장 파일 (TXT)", padding=8)
        dst_frame.pack(fill=tk.X, **pad)

        self.dst_var = tk.StringVar()
        ttk.Entry(dst_frame, textvariable=self.dst_var, font=("맑은 고딕", 9)).pack(
            side=tk.LEFT, fill=tk.X, expand=True
        )
        ttk.Button(dst_frame, text="저장 위치…", command=self._browse_dst).pack(
            side=tk.LEFT, padx=(6, 0)
        )

        # ── 진행 상태 ─────────────────────────────────────────
        prog_frame = ttk.Frame(self.root, padding=(12, 0))
        prog_frame.pack(fill=tk.X)

        self.prog_label = ttk.Label(prog_frame, text="대기 중", foreground="gray")
        self.prog_label.pack(side=tk.LEFT)

        self.prog_count = ttk.Label(prog_frame, text="", foreground="gray")
        self.prog_count.pack(side=tk.RIGHT)

        self.prog_bar = ttk.Progressbar(self.root, mode="determinate")
        self.prog_bar.pack(fill=tk.X, padx=12, pady=(2, 6))

        # ── 버튼 ─────────────────────────────────────────────
        btn_frame = ttk.Frame(self.root)
        btn_frame.pack(fill=tk.X, padx=12, pady=(0, 6))

        self.start_btn = ttk.Button(btn_frame, text="변환 시작", command=self._on_start)
        self.start_btn.pack(side=tk.LEFT)

        self.stop_btn = ttk.Button(btn_frame, text="중단", command=self._on_stop,
                                   state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(6, 0))

        self.open_btn = ttk.Button(btn_frame, text="산출물 열기", command=self._on_open_output,
                                   state=tk.DISABLED)
        self.open_btn.pack(side=tk.RIGHT)

        # ── 로그 ─────────────────────────────────────────────
        log_frame = ttk.LabelFrame(self.root, text="로그", padding=4)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=(0, 10))

        self.log = scrolledtext.ScrolledText(
            log_frame, state="disabled", height=12,
            font=("Consolas", 8), wrap="word",
            bg="#1e1e1e", fg="#d4d4d4",
        )
        self.log.pack(fill=tk.BOTH, expand=True)

    # ── 이벤트 ────────────────────────────────────────────────

    def _on_mode_change(self):
        pass  # 라디오 변경 시 추가 동작 없음

    def _browse_src(self):
        if self.mode_var.get() == "folder":
            path = filedialog.askdirectory(title="폴더 선택")
        else:
            # 체크박스 상태에 따라 파일 다이얼로그 필터를 구성.
            patterns: list[str] = []
            labels: list[str] = []
            if self.hwp_var.get():
                patterns += ["*.hwp", "*.hwpx"]
                labels.append("HWP")
            if self.pdf_var.get():
                patterns += ["*.pdf"]
                labels.append("PDF")
            if not patterns:
                patterns = ["*.hwp", "*.hwpx", "*.pdf"]
                labels = ["HWP", "PDF"]
            filter_name = "/".join(labels) + " 파일"
            path = filedialog.askopenfilename(
                title=f"{filter_name} 선택",
                filetypes=[(filter_name, " ".join(patterns)), ("모든 파일", "*.*")],
            )
        if path:
            self.src_var.set(path)

    def _browse_dst(self):
        src = self.src_var.get().strip()
        default = Path(src).stem if src else "output"
        path = filedialog.asksaveasfilename(
            title="저장 파일 이름",
            defaultextension=".txt",
            initialfile=f"{default}.txt",
            filetypes=[("텍스트 파일", "*.txt"), ("모든 파일", "*.*")],
        )
        if path:
            self.dst_var.set(path)

    def _on_start(self):
        src = self.src_var.get().strip()
        dst = self.dst_var.get().strip()

        if not src:
            messagebox.showwarning("입력 오류", "입력 경로를 선택해 주세요.")
            return
        if not dst:
            messagebox.showwarning("입력 오류", "저장 파일 경로를 지정해 주세요.")
            return
        if not Path(src).exists():
            messagebox.showwarning("경로 오류", f"경로가 존재하지 않습니다:\n{src}")
            return

        exts: set[str] = set()
        if self.hwp_var.get():
            exts |= HWP_EXTS
        if self.pdf_var.get():
            exts |= PDF_EXTS
        if not exts:
            messagebox.showwarning("포맷 선택", "HWP 또는 PDF 중 최소 1개를 선택해 주세요.")
            return

        self._stop_flag = False
        self._running   = True
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self._log_clear()

        import threading
        threading.Thread(
            target=self._run,
            args=(src, dst, self.mode_var.get() == "file", exts),
            daemon=True,
        ).start()

    def _on_stop(self):
        self._stop_flag = True
        self._ui(lambda: self.stop_btn.configure(state=tk.DISABLED))
        self._log("[중단 요청] 현재 파일 완료 후 중단합니다…")

    # ── 로그 헬퍼 ─────────────────────────────────────────────

    def _log(self, msg: str):
        def _write():
            self.log.configure(state="normal")
            self.log.insert("end", msg + "\n")
            self.log.configure(state="disabled")
            self.log.see("end")
        self.root.after(0, _write)

    def _log_clear(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _ui(self, fn):
        self.root.after(0, fn)

    def _set_progress(self, cur: int, total: int, label: str):
        def _update():
            pct = cur / total * 100 if total else 0
            self.prog_bar["value"] = pct
            self.prog_label.configure(text=label, foreground="black")
            self.prog_count.configure(text=f"{cur} / {total}")
        self.root.after(0, _update)

    # ── 변환 실행 (백그라운드 스레드) ─────────────────────────

    def _run(self, src: str, dst: str, single: bool, exts: set[str]):
        try:
            files = collect_files(src, single, exts)
        except Exception as e:
            self._log(f"[오류] 파일 수집 실패: {e}")
            self._finish()
            return

        total = len(files)
        if total == 0:
            fmt_desc = "/".join(sorted(e.lstrip(".").upper() for e in exts))
            self._log(f"[경고] {fmt_desc} 파일을 찾을 수 없습니다.")
            self._finish()
            return

        # HWP 워커는 .hwp/.hwpx 파일이 하나라도 있을 때만 spawn (한/글 띄우는
        # 비용이 크고, PDF 전용 작업에서는 불필요).
        need_hwp_worker = any(fp.suffix.lower() in HWP_EXTS for fp in files)
        hwp_count = sum(1 for fp in files if fp.suffix.lower() in HWP_EXTS)
        pdf_count = sum(1 for fp in files if fp.suffix.lower() in PDF_EXTS)
        self._log(
            f"[시작] 대상 {total}개 파일 (HWP {hwp_count} / PDF {pdf_count}) → {dst}"
        )
        t0 = time.time()

        task_q = result_q = worker = None
        if need_hwp_worker:
            task_q   = mp.Queue()
            result_q = mp.Queue()
            worker   = _spawn_worker(task_q, result_q)
            self._log(f"  HWP 워커 프로세스 시작 (PID {worker.pid})")

        ok = err = skip = 0

        try:
            with open(dst, "w", encoding="utf-8") as out:
                for i, fp in enumerate(files):
                    if self._stop_flag:
                        self._log(f"[중단] {i}개 처리 후 중단됨")
                        break

                    self._set_progress(i, total, fp.name)
                    suffix = fp.suffix.lower()

                    # PDF 는 PyMuPDF 가 \\?\ prefix 까지 처리하므로 경로 길이
                    # 제한이 없다. HWP COM 은 260자 제한이 있어 사전에 SKIP.
                    if suffix in HWP_EXTS and len(str(fp)) > 260:
                        self._log(f"  [SKIP] 경로 초과: {fp.name}")
                        skip += 1
                        continue

                    text: str | None = None
                    status = "success"
                    errmsg: str | None = None

                    if suffix in PDF_EXTS:
                        try:
                            from .pdf_parser import extract_text as pdf_extract
                            text = pdf_extract(str(fp))
                            if not text:
                                # 스캔본/이미지 PDF — 본문 없음. 에러는 아니나
                                # 출력에 빈 블록을 쓰지 않도록 skip 으로 분류.
                                status = "skip"
                                errmsg = "본문 없음(스캔본/이미지 PDF 가능성)"
                        except Exception as e:
                            status = "error"
                            errmsg = f"{type(e).__name__}: {e}"
                            text = None
                    else:
                        # HWP / HWPX → 워커 프로세스
                        task_q.put((i, str(fp), suffix))
                        try:
                            _, status, text, errmsg = result_q.get(
                                timeout=config.PARSE_TIMEOUT
                            )
                        except Empty:
                            errmsg = "타임아웃/크래시"
                            status = "error"
                            text   = None
                            try:
                                worker.kill()
                                worker.join(timeout=5)
                            except Exception:
                                pass
                            import time as _t; _t.sleep(1)
                            task_q   = mp.Queue()
                            result_q = mp.Queue()
                            worker   = _spawn_worker(task_q, result_q)
                            self._log(f"  HWP 워커 재시작 (PID {worker.pid})")

                    if status == "success" and text:
                        out.write(f"{SEPARATOR}\n")
                        out.write(f"파일: {fp.name}\n")
                        out.write(f"경로: {fp.parent}\n")
                        out.write(f"{SEPARATOR}\n")
                        out.write(text)
                        out.write("\n\n")
                        ok += 1
                        self._log(f"  [OK]  {fp.name}")
                    elif status == "error":
                        err += 1
                        self._log(f"  [ERR] {fp.name}  →  {errmsg or '알 수 없는 오류'}")
                    else:
                        skip += 1
                        reason = f"  →  {errmsg}" if errmsg else ""
                        self._log(f"  [SKIP] {fp.name}{reason}")

        except Exception as e:
            self._log(f"[오류] {e}")
        finally:
            if worker is not None:
                try:
                    task_q.put(None)
                    worker.join(timeout=10)
                except Exception:
                    pass
                try:
                    worker.kill()
                except Exception:
                    pass

        elapsed = time.time() - t0
        self._log(
            f"\n[완료] {int(elapsed//60)}분 {int(elapsed%60)}초"
            f"  성공:{ok}  오류:{err}  건너뜀:{skip}"
        )
        if ok > 0:
            self._log(f"  저장 위치: {dst}")

        self._set_progress(total, total, "완료")
        self._finish()

    def _on_open_output(self):
        dst = self.dst_var.get().strip()
        if not dst or not Path(dst).exists():
            messagebox.showwarning("파일 없음", f"파일을 찾을 수 없습니다:\n{dst}")
            return
        try:
            os.startfile(dst)
        except Exception as e:
            messagebox.showerror("열기 실패", str(e))

    def _finish(self):
        self._running = False
        def _reset():
            self.start_btn.configure(state=tk.NORMAL)
            self.stop_btn.configure(state=tk.DISABLED)
            dst = self.dst_var.get().strip()
            if dst and Path(dst).exists():
                self.open_btn.configure(state=tk.NORMAL)
        self.root.after(0, _reset)


# ═══════════════════════════════════════════════════════════════
# 진입점
# ═══════════════════════════════════════════════════════════════

def main():
    mp.freeze_support()
    root = tk.Tk()
    style = ttk.Style()
    style.configure("TLabelframe.Label", font=("맑은 고딕", 9, "bold"))
    ExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
