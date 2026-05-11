"""
앱 아이콘 — 프로그램으로 생성 (외부 파일 X, 의존성 X).

Tk PhotoImage 의 `put(color, to=rect)` 만으로 64×64 'M' 글리프를 그려
`root.iconphoto(True, img)` 에 넘김. PyInstaller onefile 환경에서도
외부 데이터 파일 의존 없이 동작.

호출자 주의:
    PhotoImage 가 GC 되면 아이콘이 사라짐. 반드시 self._icon = make_app_icon(...)
    같이 인스턴스 속성으로 들고 있을 것.
"""
from __future__ import annotations

import tkinter as tk


_BG = "#1f2937"   # slate-800
_FG = "#14b8a6"   # teal-500 (mining/data 느낌)


def make_app_icon(root: tk.Misc) -> tk.PhotoImage:
    """64×64 PhotoImage 'M' 아이콘 생성 후 반환.

    레이아웃 (64×64 그리드):
      - 배경 전체: slate
      - 좌측 다리:  x=10..18, y=12..52
      - 우측 다리:  x=46..54, y=12..52
      - 두 다리 상단에서 중앙 골(valley) (x=30..34, y=24..28) 로 내려오는
        4-step 계단형 대각선 한 쌍
    """
    img = tk.PhotoImage(master=root, width=64, height=64)
    img.put(_BG, to=(0, 0, 64, 64))

    img.put(_FG, to=(10, 12, 18, 52))   # 좌측 다리
    img.put(_FG, to=(46, 12, 54, 52))   # 우측 다리

    for i in range(4):
        y0, y1 = 12 + i * 4, 16 + i * 4
        img.put(_FG, to=(18 + i * 4, y0, 22 + i * 4, y1))   # 좌 → 중앙 계단
        img.put(_FG, to=(42 - i * 4, y0, 46 - i * 4, y1))   # 우 → 중앙 계단

    return img
