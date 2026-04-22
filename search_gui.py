"""
Step 3 — HWP 문서 검색 GUI
===========================
MariaDB에 적재된 HWP 문서를 키워드로 검색하고 클릭하면 파일을 엽니다.

단독 실행:
  python search_gui.py
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext

try:
    import pymysql
except ImportError:
    raise SystemExit("pymysql 필요: pip install pymysql")

import config

PAGE_SIZE = 200


# ═══════════════════════════════════════════════════════════════
# DB
# ═══════════════════════════════════════════════════════════════

def get_conn():
    return pymysql.connect(**config.get_db_config())


def _prepare_keywords(keyword: str, mode: str) -> list[str]:
    if mode == "phrase":
        return [keyword] if keyword else []
    return keyword.split() or []


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


def search(keyword: str, target: str, mode: str = "and",
           limit: int = PAGE_SIZE, offset: int = 0):
    keywords = _prepare_keywords(keyword, mode)
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            if keywords:
                where, params = _build_where(keywords, target, mode)
                sql = f"""
                    SELECT id, directory, filename, LEFT(body_text, 300)
                    FROM `{config.DB_TABLE}`
                    WHERE {where}
                    ORDER BY id
                    LIMIT %s OFFSET %s
                """
                cur.execute(sql, (*params, limit, offset))
            else:
                sql = f"""
                    SELECT id, directory, filename, LEFT(body_text, 300)
                    FROM `{config.DB_TABLE}`
                    ORDER BY id
                    LIMIT %s OFFSET %s
                """
                cur.execute(sql, (limit, offset))
            return cur.fetchall()
    finally:
        conn.close()


def count_results(keyword: str, target: str, mode: str = "and") -> int:
    keywords = _prepare_keywords(keyword, mode)
    conn = get_conn()
    try:
        with conn.cursor() as cur:
            if keywords:
                where, params = _build_where(keywords, target, mode)
                cur.execute(
                    f"SELECT COUNT(*) FROM `{config.DB_TABLE}` WHERE {where}", params
                )
            else:
                cur.execute(f"SELECT COUNT(*) FROM `{config.DB_TABLE}`")
            return cur.fetchone()[0]
    finally:
        conn.close()


def delete_rows(ids: list) -> int:
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
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("HWP 문서 검색기")
        self.root.geometry("1100x750")
        self.root.minsize(800, 500)

        self.results: list = []
        self._full_data: dict = {}
        self._tooltip     = None
        self._offset      = 0
        self._total       = 0
        self._last_kw     = ""
        self._last_target = "both"
        self._last_mode   = "and"

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
        self.tree.column("preview",   width=400, minwidth=100)

        vsb = ttk.Scrollbar(mid, orient=tk.VERTICAL,   command=self.tree.yview)
        hsb = ttk.Scrollbar(mid, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

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

        self.info_label = ttk.Label(
            bot,
            text="더블클릭: 파일 열기 | Del / 삭제 버튼: 선택 항목 DB에서 삭제",
            foreground="gray", font=("맑은 고딕", 9),
        )
        self.info_label.pack(side=tk.LEFT)

        self.open_btn = ttk.Button(bot, text="파일 열기",  command=self._on_open,   state=tk.DISABLED)
        self.open_btn.pack(side=tk.RIGHT)

        self.del_btn  = ttk.Button(bot, text="삭제 (Del)", command=self._on_delete, state=tk.DISABLED)
        self.del_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self.all_btn  = ttk.Button(bot, text="전체 조회",  command=self._on_load_all, state=tk.DISABLED)
        self.all_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self.more_btn = ttk.Button(bot, text=f"더보기 (+{PAGE_SIZE})", command=self._on_more, state=tk.DISABLED)
        self.more_btn.pack(side=tk.RIGHT, padx=(0, 5))

    def _bind_events(self):
        self.entry.bind("<Return>",            lambda e: self._on_search())
        self.tree.bind("<Double-1>",           lambda e: self._on_open())
        self.tree.bind("<<TreeviewSelect>>",   self._on_select)
        self.tree.bind("<Motion>",             self._on_hover)
        self.tree.bind("<Leave>",              self._hide_tooltip)
        self.tree.bind("<Delete>",             lambda e: self._on_delete())

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

        wrapped = "\n".join(text[i:i+80] for i in range(0, len(text), 80)).strip()
        if len(wrapped) > 800:
            wrapped = wrapped[:800] + "\n…"
        self._show_tooltip(event, wrapped)

    def _show_tooltip(self, event, text: str):
        self._hide_tooltip()
        tw = tk.Toplevel(self.root)
        tw.wm_overrideredirect(True)
        tw.wm_attributes("-topmost", True)

        x, y = event.x_root + 15, event.y_root + 10
        tk.Label(tw, text=text, justify=tk.LEFT,
                 background="#ffffe0", foreground="#222",
                 relief=tk.SOLID, borderwidth=1,
                 font=("맑은 고딕", 9), wraplength=600,
                 padx=6, pady=4).pack()

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
        for row in rows:
            rid, directory, filename, preview = row
            preview_full  = (preview or "").replace("\r", "").replace("\n", " ").strip()
            preview_short = preview_full[:150] + "…" if len(preview_full) > 150 else preview_full
            dir_short     = directory[:57] + "…" if len(directory) > 60 else directory
            fn_short      = filename[:50]  + "…" if len(filename) > 53  else filename

            self.tree.insert("", tk.END, iid=str(rid),
                             values=(rid, dir_short, fn_short, preview_short))
            self._full_data[str(rid)] = (rid, directory, filename, preview_full)
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

    def _on_search(self):
        kw = self.entry.get().strip()
        self.btn.configure(state=tk.DISABLED)
        self.status.configure(text="검색 중...")
        self.root.update()
        try:
            target = self.target_var.get()
            mode   = self.mode_var.get()
            self._last_kw     = kw
            self._last_target = target
            self._last_mode   = mode
            self._offset      = 0
            self._total       = count_results(kw, target, mode)
            self.results      = []
            self.tree.delete(*self.tree.get_children())
            self._full_data   = {}

            mode_label = {"and": "AND", "or": "OR", "phrase": "전체문자열"}[mode]
            self._log(f"[검색] '{kw}' | 대상: {target} | 방식: {mode_label} → {self._total:,}건")

            rows = search(kw, target, mode, offset=0)
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
            rows = search(self._last_kw, self._last_target, self._last_mode, offset=self._offset)
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
                          limit=remaining, offset=self._offset)
            self._offset += len(rows)
            self._insert_rows(rows)
            self._update_status()
        except Exception as e:
            messagebox.showerror("오류", str(e))

    # ── 선택 / 열기 ──────────────────────────────────────────────
    def _on_select(self, event=None):
        sel = self.tree.selection()
        if not sel:
            self.open_btn.configure(state=tk.DISABLED)
            self.del_btn.configure(state=tk.DISABLED)
            return

        self.del_btn.configure(state=tk.NORMAL)

        if len(sel) == 1:
            rid = sel[0]
            if rid in self._full_data:
                _, directory, filename, _ = self._full_data[rid]
                fp      = os.path.join(directory, filename)
                display = fp if len(fp) <= 90 else fp[:87] + "…"
                self.info_label.configure(text=display, foreground="black")
            self.open_btn.configure(state=tk.NORMAL)
        else:
            self.info_label.configure(text=f"{len(sel)}건 선택됨", foreground="black")
            self.open_btn.configure(state=tk.DISABLED)

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

    # ── 삭제 ─────────────────────────────────────────────────────
    def _on_delete(self):
        sel = self.tree.selection()
        if not sel:
            return

        ids, names = [], []
        for iid in sel:
            if iid in self._full_data:
                rid, _, filename, _ = self._full_data[iid]
                ids.append(rid)
                names.append(filename)

        if not ids:
            return

        preview = "\n".join(names[:10])
        if len(names) > 10:
            preview += f"\n… 외 {len(names) - 10}건"

        confirmed = messagebox.askyesno(
            "삭제 확인",
            f"DB에서 {len(ids)}건을 삭제합니다.\n\n{preview}\n\n계속하시겠습니까?",
        )
        if not confirmed:
            self._log("[삭제 취소]")
            return

        id_preview = str(ids[:10]) + ("..." if len(ids) > 10 else "")
        self._log(f"[삭제 요청] {len(ids)}건 | ID: {id_preview}")
        self.del_btn.configure(state=tk.DISABLED)

        threading.Thread(
            target=self._do_delete,
            args=(ids, sel),
            daemon=True,
        ).start()

    def _do_delete(self, ids: list, iids: tuple):
        try:
            affected = delete_rows(ids)
            self.root.after(0, lambda: self._after_delete(iids, affected))
        except Exception as e:
            self.root.after(0, lambda err=e: self._log(f"[오류] 삭제 실패: {err}"))
            self.root.after(0, lambda: self.del_btn.configure(state=tk.NORMAL))

    def _after_delete(self, iids: tuple, affected: int):
        iid_set = set(iids)
        for iid in iids:
            self._full_data.pop(iid, None)
            if self.tree.exists(iid):
                self.tree.delete(iid)

        self._total   = max(0, self._total - affected)
        self.results  = [r for r in self.results if str(r[0]) not in iid_set]
        self._offset  = max(0, self._offset - affected)

        self._update_status()
        self._log(f"[삭제 완료] {affected}건 삭제됨")
        self.del_btn.configure(state=tk.DISABLED)
        self.info_label.configure(
            text="더블클릭: 파일 열기 | Del / 삭제 버튼: 선택 항목 DB에서 삭제",
            foreground="gray",
        )


def main():
    root = tk.Tk()
    style = ttk.Style()
    style.configure("Treeview",         font=("맑은 고딕", 9), rowheight=24)
    style.configure("Treeview.Heading", font=("맑은 고딕", 9, "bold"))
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
