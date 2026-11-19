"""
src/core/analysis.py
Tổng hợp, lọc dữ liệu đọc từ bộ máy (engine.dump) và đối chiếu quy chuẩn rules/tt26.json.
"""

from __future__ import annotations

import json
import math
import os
import re
import statistics
from collections import Counter, defaultdict
from typing import Any, Dict, Iterable, List, Optional

from src.core.engine import ELEMENT_TYPES, TYPE_CODES, rgb_tuple

RULES_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
                          "rules", "tt26.json")

LINE_TYPES = {3, 4, 6, 11, 12, 14, 15, 16, 27, 36}
TEXT_TYPES = {7, 17}
CELL_TYPES = {2, 35}
DIM_TYPES = {33}


def load_rules() -> Dict[str, Any]:
    with open(RULES_PATH, encoding="utf-8") as f:
        return json.load(f)


def kind_of(t: int) -> str:
    if t in TEXT_TYPES:
        return "text"
    if t in CELL_TYPES:
        return "cell"
    if t in DIM_TYPES:
        return "dim"
    if t == 22:
        return "point"
    if t in LINE_TYPES:
        return "line"
    return "other"


def type_name(t: int) -> str:
    return ELEMENT_TYPES.get(t, f"Type{t}")


def parse_types(types: Optional[Iterable[Any]]) -> Optional[List[int]]:
    if not types:
        return None
    out = []
    for t in types:
        if isinstance(t, int) or (isinstance(t, str) and t.isdigit()):
            out.append(int(t))
        else:
            k = str(t).strip().lower()
            if k in TYPE_CODES:
                out.append(TYPE_CODES[k])
            elif k == "text":
                out += [17, 7]
            elif k == "cell":
                out += [2, 35]
            elif k == "line":
                out += [3, 4, 6, 12, 14, 16]
            else:
                raise ValueError(f"Loại đối tượng không rõ: {t}. Dùng: {sorted(TYPE_CODES)} hoặc text/cell/line")
    return out


def rgb_hex(v: int) -> Optional[str]:
    t = rgb_tuple(v)
    return None if t is None else "#%02X%02X%02X" % t


def color_group(rgb: Optional[tuple], rules: Dict[str, Any]) -> Optional[str]:
    if rgb is None:
        return None
    for g, spec in rules["color_groups"].items():
        if list(rgb) in spec["rgb"]:
            return g
    return None


def effective(e: Dict[str, Any], levels_by_code: Dict[int, Dict[str, Any]]):
    """RGB và lực nét thực tế (giải ByLevel = -1 về thuộc tính level)."""
    rgb = e.get("rgb", -1)
    w = e.get("w", -1)
    lv = levels_by_code.get(e.get("lv"))
    if (rgb is None or rgb < 0) and lv:
        rgb = lv.get("rgb", -1)
    if (w is None or w < 0) and lv:
        w = lv.get("weight", -1)
    return rgb_tuple(rgb), w


def seg_length(v: List[List[float]]) -> float:
    return sum(math.dist(v[i], v[i + 1]) for i in range(len(v) - 1)) if v and len(v) > 1 else 0.0


# ------------------------------------------------------------------ summary ---
def summarize(data: Dict[str, Any], rules: Dict[str, Any], samples: int = 4) -> Dict[str, Any]:
    els = data["elements"]
    hdr = data["header"]
    levels_by_code = {l["code"]: l for l in hdr.get("levels", [])}
    per = defaultdict(lambda: {"count": 0, "types": Counter(), "colors": Counter(), "weights": Counter(),
                               "styles": Counter(), "texts": [], "cells": Counter(), "lengths": []})
    xs, ys = [], []
    for e in els:
        p = per[e["lv"]]
        p["count"] += 1
        p["types"][type_name(e["t"])] += 1
        rgb, w = effective(e, levels_by_code)
        p["colors"][f'{e.get("c")}{"=" + ("#%02X%02X%02X" % rgb) if rgb else ""}'] += 1
        p["weights"][w] += 1
        if e.get("ls") != "?":
            p["styles"][e.get("ls", "")] += 1
        if "tx" in e and len(p["texts"]) < samples:
            p["texts"].append(e["tx"][:60])
        if "cell" in e:
            p["cells"][e["cell"]] += 1
        if "v" in e and e["t"] in (3, 4):
            p["lengths"].append(seg_length(e["v"]))
        r = e.get("r")
        if r and abs(r[0]) < 1e8 and r[0] > -1e6:  # bỏ phần tử điều khiển pattern có range ảo
            xs += [r[0], r[2]]
            ys += [r[1], r[3]]

    pl21 = rules["levels"]
    out_levels = []
    for code in sorted(per):
        p = per[code]
        rule = pl21.get(str(code))
        lvinfo = levels_by_code.get(code, {})
        item = {
            "level": code,
            "file_name": lvinfo.get("name") if code >= 0 else "(không có level - cell/phần tử điều khiển)",
            "pl21": rule["name"] if rule else (rules.get("removed_levels", {}).get(str(code))
                                               or ("KHÔNG CÓ TRONG PL21" if code > 0 else None)),
            "count": p["count"],
            "types": dict(p["types"]),
            "colors(index=rgb)": dict(p["colors"].most_common(4)),
            "weights": dict(p["weights"].most_common(3)),
            "styles": dict(p["styles"].most_common(3)),
        }
        if p["texts"]:
            item["text_samples"] = p["texts"]
        if p["cells"]:
            item["cells"] = dict(p["cells"].most_common(8))
        if p["lengths"]:
            L = p["lengths"]
            item["line_length_m"] = {"n": len(L), "min": round(min(L), 3), "median": round(statistics.median(L), 3),
                                     "max": round(max(L), 3)}
        out_levels.append(item)
    return {
        "file": hdr.get("file"),
        "model": hdr.get("model"),
        "opened_read_only_in_background": hdr.get("external", False),
        "engine": data.get("engine"),
        "read_seconds": data.get("seconds"),
        "total_elements": len(els),
        "extent": [round(min(xs), 3), round(min(ys), 3), round(max(xs), 3), round(max(ys), 3)] if xs else None,
        "levels_used": len(per),
        "by_level": out_levels,
    }


# -------------------------------------------------------------------- query ---
def compact(e: Dict[str, Any], geometry: bool = True, vmax: int = 50) -> Dict[str, Any]:
    d = {"id": e["id"], "type": type_name(e["t"]), "level": e["lv"], "color": e.get("c"),
         "rgb": rgb_hex(e.get("rgb", -1)), "weight": e.get("w"), "style": e.get("ls"),
         "range": e.get("r")}
    for k_src, k_dst in (("tx", "text"), ("cell", "cell"), ("o", "origin"), ("h", "height"), ("rad", "radius")):
        if k_src in e:
            d[k_dst] = e[k_src]
    if geometry and "v" in e:
        v = e["v"]
        d["vertices"] = v if len(v) <= vmax else v[:vmax] + [f"... (+{len(v) - vmax} đỉnh)"]
        d["n_vertices"] = len(v)
        if len(v) > 1:
            d["length"] = round(seg_length(v), 3)
            if e["t"] in (6, 14) or (len(v) > 2 and v[0] == v[-1]):
                d["closed"] = True
                d["area"] = round(abs(_shoelace(v)), 3)
    if e.get("gg"):
        d["graphic_group"] = e["gg"]
    return d


def _shoelace(v: List[List[float]]) -> float:
    s = 0.0
    for i in range(len(v) - 1):
        s += v[i][0] * v[i + 1][1] - v[i + 1][0] * v[i][1]
    return s / 2.0


def filter_elements(els: List[Dict[str, Any]], text: Optional[str] = None, text_regex: Optional[str] = None,
                    cell: Optional[str] = None, color: Optional[int] = None) -> List[Dict[str, Any]]:
    out = els
    if text is not None:
        tl = text.lower()
        out = [e for e in out if "tx" in e and tl in e["tx"].lower()]
    if text_regex:
        rx = re.compile(text_regex)
        out = [e for e in out if "tx" in e and rx.search(e["tx"])]
    if cell:
        cl = cell.lower()
        out = [e for e in out if e.get("cell", "").lower() == cl]
    if color is not None:
        out = [e for e in out if e.get("c") == color]
    return out
