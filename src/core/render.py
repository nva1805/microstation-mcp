"""
src/core/render.py
Dựng ảnh bản vẽ từ dữ liệu engine.dump (không phụ thuộc cửa sổ MicroStation).

Lý do: khi cửa sổ MicroStation bị che, nó không vẽ lại; PrintWindow trả ảnh cũ hoặc đen.
Dựng từ dữ liệu luôn đúng vùng, chạy cả khi MicroStation thu nhỏ, và tô đậm được đối tượng cần chỉ.
Không gồm file tham chiếu (reference) - chỉ dữ liệu của chính file.
"""

from __future__ import annotations

import io
import math
import os
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

from PIL import Image, ImageDraw, ImageFont

from src.core.engine import rgb_tuple

_FONT_PATHS = [r"C:\Windows\Fonts\arial.ttf", r"C:\Windows\Fonts\tahoma.ttf", r"C:\Windows\Fonts\segoeui.ttf"]
_FONT_FILE = next((p for p in _FONT_PATHS if os.path.exists(p)), None)
_fonts: Dict[int, ImageFont.ImageFont] = {}

HILITE = (230, 0, 120)


def _font(px: int):
    px = max(6, min(px, 200))
    if px not in _fonts:
        _fonts[px] = ImageFont.truetype(_FONT_FILE, px) if _FONT_FILE else ImageFont.load_default()
    return _fonts[px]


def _ink(e) -> Tuple[int, int, int]:
    c = rgb_tuple(e.get("rgb", -1))
    if c is None:
        return (0, 0, 0)
    if sum(c) > 690:  # trắng (chỉ số 0) -> in ra đen trên nền trắng
        return (0, 0, 0)
    if c[0] > 200 and c[1] > 200 and c[2] < 80:  # vàng -> sẫm lại cho dễ nhìn trên nền trắng
        return (170, 140, 0)
    return c


def _dashed(draw, pts, fill, width, dash=(10, 6)):
    on, off = dash
    for (x1, y1), (x2, y2) in zip(pts, pts[1:]):
        seg = math.hypot(x2 - x1, y2 - y1)
        if seg == 0:
            continue
        ux, uy = (x2 - x1) / seg, (y2 - y1) / seg
        d = 0.0
        while d < seg:
            e = min(d + on, seg)
            draw.line([(x1 + ux * d, y1 + uy * d), (x1 + ux * e, y1 + uy * e)], fill=fill, width=width)
            d = e + off


def render(elements: Sequence[dict], bbox: Optional[Sequence[float]] = None, width: int = 1400,
           highlight: Iterable[str] = (), fmt: str = "png", title: Optional[str] = None) -> bytes:
    els = [e for e in elements if e.get("r") and e["r"][0] > -1e6]
    if not bbox:
        if not els:
            bbox = [0, 0, 1, 1]
        else:
            bbox = [min(e["r"][0] for e in els), min(e["r"][1] for e in els),
                    max(e["r"][2] for e in els), max(e["r"][3] for e in els)]
    x0, y0, x1, y1 = bbox
    w_m, h_m = max(x1 - x0, 1e-6), max(y1 - y0, 1e-6)
    height = max(200, min(int(width * h_m / w_m), 2000))
    s = min(width / w_m, height / h_m)
    ox = (width - w_m * s) / 2
    oy = (height - h_m * s) / 2

    def P(p):
        return (ox + (p[0] - x0) * s, height - (oy + (p[1] - y0) * s))

    hi = set(highlight)
    img = Image.new("RGB", (width, height), "white")
    dr = ImageDraw.Draw(img)
    vis = [e for e in els if not (e["r"][2] < x0 or e["r"][0] > x1 or e["r"][3] < y0 or e["r"][1] > y1)]

    # 1) đường, vùng, cung
    for e in vis:
        is_hi = e["id"] in hi
        col = HILITE if is_hi else _ink(e)
        wpx = 3 if is_hi else (2 if (e.get("w") or 0) >= 2 else 1)
        if "v" in e and len(e["v"]) >= 2:
            pts = [P(p) for p in e["v"]]
            style = str(e.get("ls", "0"))
            if style not in ("0", "", "?", "ByLevel") and not is_hi:
                _dashed(dr, pts, col, wpx)
            else:
                dr.line(pts, fill=col, width=wpx)
        elif e["t"] in (15, 16) and "o" in e and "rad" in e:
            cx, cy = P(e["o"])
            r = e["rad"] * s
            if r >= 1:
                dr.ellipse([cx - r, cy - r, cx + r, cy + r], outline=col, width=wpx)
    # 2) cell: dấu chữ thập nhỏ
    for e in vis:
        if "cell" in e and "o" in e:
            cx, cy = P(e["o"])
            col = HILITE if e["id"] in hi else (90, 90, 90)
            dr.line([(cx - 4, cy), (cx + 4, cy)], fill=col)
            dr.line([(cx, cy - 4), (cx, cy + 4)], fill=col)
    # 3) chữ (chỉ khi đủ lớn để đọc)
    for e in vis:
        if "tx" in e and "o" in e:
            px = int((e.get("h") or 2.0) * s)
            if px < 7 and e["id"] not in hi:
                continue
            col = HILITE if e["id"] in hi else _ink(e)
            x, y = P(e["o"])
            dr.text((x, y), e["tx"].split("\n")[0][:40], fill=col, font=_font(max(px, 9)), anchor="lb")
    # 4) thước tỉ lệ + tiêu đề
    bar_m = 10 ** math.floor(math.log10(max(w_m / 5, 1e-3)))
    for k in (5, 2):
        if bar_m * k * s < width / 4:
            bar_m *= k
            break
    bx, by = 15, height - 15
    dr.line([(bx, by), (bx + bar_m * s, by)], fill=(0, 0, 0), width=3)
    label = f"{bar_m:g} m" + (f"   {title}" if title else "")
    dr.text((bx, by - 6), label, fill=(0, 0, 0), font=_font(13), anchor="lb")
    buf = io.BytesIO()
    if fmt == "jpeg":
        img.save(buf, "JPEG", quality=85, optimize=True)
    else:
        img.save(buf, "PNG", optimize=True)
    return buf.getvalue()
