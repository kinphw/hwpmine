"""
스캔할 드라이브 선택 다이얼로그.

Win32 API 로 현재 마운트된 드라이브를 열거하여 체크박스 목록으로 표시한다.
호출 측은 `pick_drives(defaults)` 로 부르고 사용자가 확정한 드라이브 리스트를
받거나, 취소 시 None 을 받는다.
"""
from __future__ import annotations

import ctypes
import string
import tkinter as tk
from ctypes import wintypes
from tkinter import ttk
from typing import Optional

from .icon import make_app_icon


# ── Win32 ─────────────────────────────────────────────────────
_DRIVE_TYPE_LABEL = {
    0: "알수없음",
    1: "없음",
    2: "이동식",
    3: "고정",
    4: "네트워크",
    5: "CD-ROM",
    6: "램디스크",
}

_kernel32 = ctypes.windll.kernel32
_kernel32.GetLogicalDrives.restype  = wintypes.DWORD
_kernel32.GetDriveTypeW.restype     = wintypes.UINT
_kernel32.GetDriveTypeW.argtypes    = [wintypes.LPCWSTR]
_kernel32.GetVolumeInformationW.argtypes = [
    wintypes.LPCWSTR,                       # lpRootPathName
    wintypes.LPWSTR,                        # lpVolumeNameBuffer
    wintypes.DWORD,                         # nVolumeNameSize
    ctypes.POINTER(wintypes.DWORD),         # lpVolumeSerialNumber
    ctypes.POINTER(wintypes.DWORD),         # lpMaximumComponentLength
    ctypes.POINTER(wintypes.DWORD),         # lpFileSystemFlags
    wintypes.LPWSTR,                        # lpFileSystemNameBuffer
    wintypes.DWORD,                         # nFileSystemNameSize
]


def list_drives() -> list[tuple[str, str, str]]:
    """현재 마운트된 드라이브 목록 (외부 공개 API).

    반환: [(root, label, drive_type_name)] — 예) ("C:\\", "Windows", "고정")
    """
    bitmask = _kernel32.GetLogicalDrives()
    drives: list[tuple[str, str, str]] = []
    for i, letter in enumerate(string.ascii_uppercase):
        if not (bitmask & (1 << i)):
            continue
        root = f"{letter}:\\"
        dtype = _DRIVE_TYPE_LABEL.get(_kernel32.GetDriveTypeW(root), "알수없음")

        # 네트워크/CD-ROM 등은 라벨 조회가 느리거나 실패할 수 있어 try
        label = ""
        try:
            buf = ctypes.create_unicode_buffer(261)
            if _kernel32.GetVolumeInformationW(
                root, buf, 261, None, None, None, None, 0,
            ):
                label = buf.value
        except OSError:
            pass

        drives.append((root, label, dtype))
    return drives


# 하위 호환 별칭 — 다른 곳에서 _list_drives 를 import 했다면.
_list_drives = list_drives


# ── 다이얼로그 ────────────────────────────────────────────────
def pick_drives(defaults: list[str], parent: Optional[tk.Misc] = None) -> Optional[list[str]]:
    """체크박스로 스캔 대상 드라이브를 선택받고, 확정된 리스트 또는 None 반환.

    Args:
        defaults: 초기 체크 상태로 표시할 드라이브 루트 목록 (예: ["C:\\", "D:\\"]).
        parent: 부모 윈도우. 주어지면 Toplevel(모달) 로 띄우고, None 이면
                새 tk.Tk() 를 만들어 단독 다이얼로그로 동작 (콘솔 호출용).
    """
    drives = list_drives()
    default_set = {d.upper().rstrip("\\") + "\\" for d in defaults}

    embedded = parent is not None
    if embedded:
        root = tk.Toplevel(parent)
        root.transient(parent.winfo_toplevel())
    else:
        root = tk.Tk()
    root.title("스캔할 드라이브 선택")
    root.resizable(False, False)

    try:
        icon = make_app_icon(root)
        root.iconphoto(True, icon)
        root._icon_ref = icon  # GC 방지
    except tk.TclError:
        pass

    result: dict[str, Optional[list[str]]] = {"value": None}

    body = ttk.Frame(root, padding=16)
    body.pack(fill=tk.BOTH, expand=True)

    ttk.Label(
        body, text="스캔할 드라이브를 선택하세요 (.hwp / .hwpx 검색)",
        font=("맑은 고딕", 10, "bold"),
    ).pack(anchor="w", pady=(0, 10))

    list_frame = ttk.Frame(body)
    list_frame.pack(fill=tk.BOTH, expand=True)

    vars_: list[tuple[tk.BooleanVar, str]] = []
    if not drives:
        ttk.Label(list_frame, text="(마운트된 드라이브가 없습니다)",
                  foreground="gray").pack(anchor="w")
    else:
        for path, label, dtype in drives:
            init = path.upper() in default_set
            v = tk.BooleanVar(value=init)
            text = f"{path}    {label or '(라벨 없음)'}    [{dtype}]"
            ttk.Checkbutton(list_frame, text=text, variable=v).pack(anchor="w")
            vars_.append((v, path))

    # 전체 선택/해제
    quick = ttk.Frame(body)
    quick.pack(fill=tk.X, pady=(10, 0))

    def _set_all(value: bool):
        for v, _ in vars_:
            v.set(value)

    ttk.Button(quick, text="전체 선택", command=lambda: _set_all(True)).pack(side=tk.LEFT)
    ttk.Button(quick, text="전체 해제", command=lambda: _set_all(False)).pack(side=tk.LEFT, padx=(6, 0))

    # 하단 버튼
    bot = ttk.Frame(body)
    bot.pack(fill=tk.X, pady=(14, 0))

    def _on_ok():
        sel = [p for v, p in vars_ if v.get()]
        result["value"] = sel
        root.destroy()

    def _on_cancel():
        result["value"] = None
        root.destroy()

    ttk.Button(bot, text="취소", command=_on_cancel, width=10).pack(side=tk.RIGHT)
    ttk.Button(bot, text="스캔 시작", command=_on_ok, width=12).pack(side=tk.RIGHT, padx=(0, 6))

    root.bind("<Escape>", lambda _e: _on_cancel())
    root.bind("<Return>", lambda _e: _on_ok())
    root.protocol("WM_DELETE_WINDOW", _on_cancel)

    # 화면 중앙 배치
    root.update_idletasks()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    ww, wh = root.winfo_reqwidth(), root.winfo_reqheight()
    root.geometry(f"+{(sw - ww) // 2}+{(sh - wh) // 3}")

    if embedded:
        # 모달 다이얼로그처럼 동작 — 부모 mainloop 안에서 wait_window 로 블록.
        root.grab_set()
        root.wait_window()
    else:
        root.mainloop()
    return result["value"]
