"""
HWP COM 단독 테스트 스크립트 — 개발환경 진단용.

사용법:
    python test_hwp_com.py <hwp파일경로>
    python test_hwp_com.py <hwp파일경로> --clear-cache    # gen_py 캐시 삭제 후 재시도
    python test_hwp_com.py <hwp파일경로> --dispatch       # EnsureDispatch 대신 Dispatch 사용
    python test_hwp_com.py <hwp파일경로> --visible        # 한/글 창 보이게
    python test_hwp_com.py <hwp파일경로> --popup-loop     # 보안/경고 다이얼로그 자동 dismiss 스레드 가동

각 단계 소요시간을 출력해 어디서 막히는지 확인.
워커/큐/MariaDB 없이 순수 COM 만 검증.
"""
from __future__ import annotations

import argparse
import shutil
import sys
import threading
import time
from pathlib import Path


def start_popup_loop():
    """inserter.worker_main 과 동일한 팝업 dismiss 스레드."""
    try:
        import win32gui
    except ImportError:
        print("    (win32gui 없음 — popup loop 비활성)")
        return

    BTNS = ["접근 허용(&A)", "접근 허용", "확인(&O)", "확인", "OK",
            "아니오(&N)", "예(&Y)", "취소(&C)", "취소",
            "저장(&Y)", "저장"]

    seen = set()

    def _loop():
        while True:
            try:
                def _on(h, _):
                    if not win32gui.IsWindowVisible(h):
                        return
                    try:
                        title = win32gui.GetWindowText(h)
                    except Exception:
                        title = ""
                    if title and title not in seen:
                        seen.add(title)
                        print(f"\n    [popup] 다이얼로그 감지: {title!r}", flush=True)

                    def _c(c, _):
                        try:
                            if win32gui.GetClassName(c) != "Button":
                                return
                            t = win32gui.GetWindowText(c)
                            if any(b in t for b in BTNS):
                                print(f"    [popup] 버튼 클릭: {t!r}", flush=True)
                                win32gui.SendMessage(c, 0xF5, 0, 0)
                        except Exception:
                            pass
                    try:
                        win32gui.EnumChildWindows(h, _c, None)
                    except Exception:
                        pass
                win32gui.EnumWindows(_on, None)
            except Exception:
                pass
            time.sleep(0.3)

    threading.Thread(target=_loop, daemon=True).start()


def clear_gencache() -> None:
    """%TEMP%\\gen_py 캐시 디렉터리 삭제."""
    import tempfile
    gen = Path(tempfile.gettempdir()) / "gen_py"
    if gen.exists():
        print(f"  [gencache] 삭제: {gen}")
        shutil.rmtree(gen, ignore_errors=True)
    else:
        print(f"  [gencache] 없음: {gen}")


def step(label: str, fn):
    t0 = time.time()
    print(f"  [{label}] 시작…", flush=True)
    try:
        rv = fn()
        dt = time.time() - t0
        print(f"  [{label}] OK ({dt:.2f}s)", flush=True)
        return rv
    except Exception as e:
        dt = time.time() - t0
        print(f"  [{label}] FAIL ({dt:.2f}s): {type(e).__name__}: {e}", flush=True)
        raise


def main() -> int:
    ap = argparse.ArgumentParser(description="HWP COM 단독 테스트")
    ap.add_argument("filepath", help="테스트할 .hwp / .hwpx 경로")
    ap.add_argument("--clear-cache", action="store_true",
                    help="실행 전 win32com gen_py 캐시 삭제")
    ap.add_argument("--dispatch", action="store_true",
                    help="EnsureDispatch 대신 Dispatch 사용 (frozen-exe 호환 모드)")
    ap.add_argument("--dispatchex", action="store_true",
                    help="DispatchEx 사용 — 신규 한/글 프로세스 강제 생성 (ROT 어태치 회피)")
    ap.add_argument("--visible", action="store_true",
                    help="한/글 창을 보이게 (디버깅용)")
    ap.add_argument("--popup-loop", action="store_true",
                    help="보안/경고 다이얼로그를 자동 dismiss (inserter 와 동일)")
    args = ap.parse_args()

    fp = Path(args.filepath).expanduser().resolve()
    if not fp.exists():
        print(f"파일이 존재하지 않습니다: {fp}")
        return 1

    print("=" * 60)
    print(f"  대상: {fp}")
    print(f"  크기: {fp.stat().st_size:,} bytes")
    print(f"  옵션: clear_cache={args.clear_cache} dispatch={args.dispatch} visible={args.visible}")
    print("=" * 60)

    if args.clear_cache:
        clear_gencache()

    if args.popup_loop:
        print("  [popup loop] 시작")
        start_popup_loop()

    import win32com.client as win32

    def _make_com():
        if args.dispatchex:
            return win32.DispatchEx("HwpFrame.HwpObject")
        if args.dispatch:
            return win32.Dispatch("HwpFrame.HwpObject")
        return win32.gencache.EnsureDispatch("HwpFrame.HwpObject")

    com = step("COM dispatch", _make_com)

    def _register():
        try:
            com.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
            return True
        except Exception as e:
            print(f"    (RegisterModule 실패 — 무시: {e})")
            return False

    step("RegisterModule", _register)

    if not args.visible:
        def _hide():
            com.XHwpWindows.Item(0).Visible = False
        step("Hide window", _hide)

    def _msgbox_off():
        try:
            com.SetMessageBoxMode(0x10000)
        except Exception as e:
            print(f"    (SetMessageBoxMode 실패 — 무시: {e})")
    step("SetMessageBoxMode(0x10000)", _msgbox_off)

    def _open():
        # HWP Open(Path, Format, Arg) — late-binding(Dispatch)에선 디폴트 미적용이라
        # 3 인자를 명시. "forceopen:true" 로 손상/경고 파일도 열도록 함.
        com.Open(str(fp), "", "forceopen:true")
    step("Open", _open)

    def _get_text():
        return com.GetTextFile("TEXT", "")
    raw = step("GetTextFile", _get_text)

    if raw:
        # 미리보기만 출력
        preview = raw[:200].replace("\r", " ").replace("\n", " ")
        print(f"  [본문] {len(raw):,}자 / 미리보기: {preview}…")
    else:
        print("  [본문] 빈 텍스트")

    def _set_unmod():
        try:
            com.XHwpDocuments.Item(0).SetModified(False)
        except Exception:
            pass
    step("SetModified(False)", _set_unmod)

    def _close():
        try:
            com.Run("FileClose")
        except Exception as e:
            print(f"    (FileClose 실패 — 무시: {e})")
    step("FileClose", _close)

    print()
    print("=" * 60)
    print("  완료 — COM 파이프라인 정상 동작")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    sys.exit(main())
