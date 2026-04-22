"""
HWP / HWPX → TXT 변환기
=========================
파일 1개 또는 폴더 내 모든 HWP/HWPX를 파싱하여
하나의 TXT 파일로 추출합니다.

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

try:
    from inserter import worker_main, _kill_hwp
except ImportError:
    raise SystemExit("inserter.py를 이 스크립트와 같은 폴더에 두세요.")

try:
    import config
except ImportError:
    raise SystemExit("config.py를 이 스크립트와 같은 폴더에 두세요.")

TARGET_EXT = {".hwp", ".hwpx"}
SEPARATOR  = "=" * 80


# ═══════════════════════════════════════════════════════════════
# 파일 수집
# ═══════════════════════════════════════════════════════════════

def collect_files(path: str, single: bool) -> list[Path]:
    p = Path(path)
    if single:
        return [p] if p.suffix.lower() in TARGET_EXT else []
    return sorted(
        f for f in p.rglob("*") if f.suffix.lower() in TARGET_EXT
    )


# ═══════════════════════════════════════════════════════════════
# 워커
# ═══════════════════════════════════════════════════════════════

def _spawn_worker(task_q, result_q):
    w = mp.Process(target=worker_main, args=(task_q, result_q), daemon=True)
    w.start()
    return w


# ═══════════════════════════════════════════════════════════════
# GUI
# ═══════════════════════════════════════════════════════════════

class ExtractorApp:
    def __init__(self, root: tk.Tk):
        self.root  = root
        self.root.title("HWP / HWPX → TXT 변환기")
        self.root.geometry("720x560")
        self.root.minsize(600, 480)
        self.root.resizable(True, True)

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
            path = filedialog.askopenfilename(
                title="HWP/HWPX 파일 선택",
                filetypes=[("HWP 파일", "*.hwp *.hwpx"), ("모든 파일", "*.*")],
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

        self._stop_flag = False
        self._running   = True
        self.start_btn.configure(state=tk.DISABLED)
        self.stop_btn.configure(state=tk.NORMAL)
        self._log_clear()

        import threading
        threading.Thread(
            target=self._run,
            args=(src, dst, self.mode_var.get() == "file"),
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

    def _run(self, src: str, dst: str, single: bool):
        try:
            files = collect_files(src, single)
        except Exception as e:
            self._log(f"[오류] 파일 수집 실패: {e}")
            self._finish()
            return

        total = len(files)
        if total == 0:
            self._log("[경고] HWP/HWPX 파일을 찾을 수 없습니다.")
            self._finish()
            return

        self._log(f"[시작] 대상 {total}개 파일 → {dst}")
        t0 = time.time()

        task_q   = mp.Queue()
        result_q = mp.Queue()
        worker   = _spawn_worker(task_q, result_q)
        self._log(f"  워커 프로세스 시작 (PID {worker.pid})")

        ok = err = skip = 0

        try:
            with open(dst, "w", encoding="utf-8") as out:
                for i, fp in enumerate(files):
                    if self._stop_flag:
                        self._log(f"[중단] {i}개 처리 후 중단됨")
                        break

                    self._set_progress(i, total, fp.name)

                    if len(str(fp)) > 260:
                        self._log(f"  [SKIP] 경로 초과: {fp.name}")
                        skip += 1
                        continue

                    task_q.put((i, str(fp), fp.suffix.lower()))

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
                        _kill_hwp()
                        import time as _t; _t.sleep(1)
                        task_q   = mp.Queue()
                        result_q = mp.Queue()
                        worker   = _spawn_worker(task_q, result_q)
                        self._log(f"  워커 재시작 (PID {worker.pid})")

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
                        self._log(f"  [SKIP] {fp.name}")

        except Exception as e:
            self._log(f"[오류] {e}")
        finally:
            try:
                task_q.put(None)
                worker.join(timeout=10)
            except Exception:
                pass
            try:
                worker.kill()
            except Exception:
                pass
            _kill_hwp()

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
