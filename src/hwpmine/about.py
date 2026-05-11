"""About 다이얼로그 — search_gui / extractor_gui 공용."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from . import __app_name__, __author__, __tagline__, __version__


def show_about(parent: tk.Misc) -> None:
    """? 버튼용 — 버전·작성자·태그라인을 표시하는 모달 다이얼로그."""
    win = tk.Toplevel(parent)
    win.title(f"About {__app_name__}")
    win.transient(parent.winfo_toplevel())
    win.grab_set()
    win.resizable(False, False)

    body = ttk.Frame(win, padding=24)
    body.pack(fill=tk.BOTH, expand=True)

    ttk.Label(body, text=__app_name__, font=("", 20, "bold")).pack(anchor="w")
    ttk.Label(body, text=f"v{__version__}", foreground="#888").pack(anchor="w", pady=(0, 14))
    ttk.Label(body, text=f"Author: {__author__}").pack(anchor="w")
    ttk.Label(
        body, text=__tagline__,
        foreground="#555", font=("", 10, "italic"),
    ).pack(anchor="w", pady=(8, 16))

    ttk.Button(body, text="확인", command=win.destroy, width=10).pack(anchor="e")

    # 부모 윈도우 중앙 정렬
    try:
        top = parent.winfo_toplevel()
        top.update_idletasks()
        win.update_idletasks()
        px, py = top.winfo_rootx(), top.winfo_rooty()
        pw, ph = top.winfo_width(), top.winfo_height()
        ww, wh = win.winfo_reqwidth(), win.winfo_reqheight()
        x = px + (pw - ww) // 2
        y = py + (ph - wh) // 3
        win.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    except Exception:
        pass

    win.bind("<Escape>", lambda _e: win.destroy())
    win.bind("<Return>", lambda _e: win.destroy())
