"""
src/core/screenshot.py
Chụp khung bản vẽ MicroStation V8i.

Một lần PrintWindow cửa sổ chính (chụp được cả khi bị che), rồi cắt đúng vùng khung nhìn
"VpWindow <n>" - không lẫn menu, thanh công cụ, Project Explorer. PNG hoặc JPEG, tự thu nhỏ.
"""

from __future__ import annotations

import ctypes
import io
from typing import Optional, Tuple

import psutil
import win32gui
import win32process
import win32ui
from PIL import Image

_user32 = ctypes.windll.user32
try:  # tọa độ thật trên màn hình scale 125%/150%
    _user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
except Exception:  # noqa: BLE001
    _user32.SetProcessDPIAware()


def main_window() -> Optional[int]:
    """Cửa sổ cấp cao lớn nhất của tiến trình ustation.exe."""
    pids = {p.pid for p in psutil.process_iter(["name"]) if (p.info["name"] or "").lower() == "ustation.exe"}
    found = []

    def visit(h, _):
        if win32gui.IsWindowVisible(h) and not win32gui.GetParent(h):
            if win32process.GetWindowThreadProcessId(h)[1] in pids:
                l, t, r, b = win32gui.GetWindowRect(h)
                found.append(((r - l) * (b - t), h))
        return True

    win32gui.EnumWindows(visit, None)
    return max(found)[1] if found else None


def viewport(main: int, view: int) -> Optional[int]:
    hit = []
    win32gui.EnumChildWindows(
        main, lambda h, _: hit.append(h) if win32gui.GetWindowText(h) == f"VpWindow {view}"
        and win32gui.IsWindowVisible(h) else None, None)
    return hit[0] if hit else None


def _grab(hwnd: int) -> Image.Image:
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r - l, b - t
    src = win32gui.GetWindowDC(hwnd)
    dc = win32ui.CreateDCFromHandle(src)
    mem = dc.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    try:
        bmp.CreateCompatibleBitmap(dc, w, h)
        mem.SelectObject(bmp)
        _user32.PrintWindow(hwnd, mem.GetSafeHdc(), 2)  # PW_RENDERFULLCONTENT
        return Image.frombuffer("RGB", (w, h), bmp.GetBitmapBits(True), "raw", "BGRX", 0, 1)
    finally:
        win32gui.DeleteObject(bmp.GetHandle())
        mem.DeleteDC()
        dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, src)


def capture(view: int = 1, full_window: bool = False, fmt: str = "png",
            max_width: int = 1600) -> Tuple[Optional[bytes], str]:
    """Trả (bytes, mô tả). bytes=None kèm lý do nếu không chụp được."""
    main = main_window()
    if not main:
        return None, "Không tìm thấy cửa sổ MicroStation"
    if win32gui.IsIconic(main):
        return None, "Cửa sổ MicroStation đang thu nhỏ - hãy mở lại"
    img = _grab(main)
    what = "toàn cửa sổ"
    vp = None if full_window else viewport(main, view)
    if vp:
        ml, mt, _, _ = win32gui.GetWindowRect(main)
        l, t, r, b = win32gui.GetWindowRect(vp)
        img = img.crop((l - ml, t - mt, r - ml, b - mt))
        what = f"khung nhìn {view}"
    if img.width > max_width:
        img = img.resize((max_width, round(img.height * max_width / img.width)), Image.LANCZOS)
    buf = io.BytesIO()
    if fmt == "jpeg":
        img.save(buf, "JPEG", quality=82, optimize=True)
    else:
        img.save(buf, "PNG", optimize=True)
    return buf.getvalue(), what
