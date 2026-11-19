"""
src/core/kcs.py
Kiểm tra KCS bản đồ địa chính theo TT 26/2024 (PL21, PL22) trên dữ liệu engine.dump.

Mỗi lỗi (finding) có: check, severity (error|warning|info), message, level, ids, xy.
Mọi tọa độ là mét (VN-2000) theo file.
"""

from __future__ import annotations

import csv
import math
import os
import re
import time
from collections import Counter, defaultdict
from typing import Any, Dict, List, Optional, Tuple

from shapely import STRtree, polygonize_full, unary_union
from shapely.geometry import LineString, MultiLineString, Point, Polygon

from src.core.analysis import effective, kind_of, type_name, color_group

ALL_CHECKS = ("levels", "symbology", "topology", "parcels")


def _f(check, sev, msg, level=None, ids=None, xy=None, **extra) -> Dict[str, Any]:
    d = {"check": check, "severity": sev, "message": msg}
    if level is not None:
        d["level"] = level
    if ids:
        d["ids"] = list(ids)[:20]
        if len(ids) > 20:
            d["ids_more"] = len(ids) - 20
    if xy is not None:
        d["xy"] = [round(xy[0], 3), round(xy[1], 3)]
    d.update(extra)
    return d


# ---------------------------------------------------------------- levels ---
def check_levels(els, rules) -> List[Dict[str, Any]]:
    out = []
    pl21 = rules["levels"]
    removed = rules.get("removed_levels", {})
    by_lv = defaultdict(list)
    for e in els:
        by_lv[e["lv"]].append(e)
    for lv, items in sorted(by_lv.items()):
        if lv < 0:
            continue  # cell/phần tử không gắn level
        ids = [e["id"] for e in items]
        if str(lv) in removed:
            out.append(_f("levels", "error", f"Level {lv}: {removed[str(lv)]} - còn {len(items)} đối tượng", lv, ids))
            continue
        rule = pl21.get(str(lv))
        if not rule:
            types = Counter(type_name(e["t"]) for e in items)
            out.append(_f("levels", "error",
                          f"Level {lv} không có trong bảng phân lớp PL21 - {len(items)} đối tượng {dict(types)}",
                          lv, ids))
            continue
        bad = defaultdict(list)
        for e in items:
            k = kind_of(e["t"])
            if k not in rule["kinds"] and k != "other":
                bad[k].append(e["id"])
        for k, bids in bad.items():
            out.append(_f("levels", "warning",
                          f"Level {lv} ({rule['name']}) có {len(bids)} đối tượng dạng '{k}', "
                          f"PL21 chỉ dự kiến {rule['kinds']}", lv, bids))
    return out


# ------------------------------------------------------------- symbology ---
def check_symbology(els, rules, levels_by_code) -> List[Dict[str, Any]]:
    out = []
    pl21 = rules["levels"]
    groups = rules["color_groups"]
    bad = defaultdict(list)
    solid_house = []
    for e in els:
        rule = pl21.get(str(e["lv"]))
        if not rule:
            continue
        rgb, _w = effective(e, levels_by_code)
        g = color_group(rgb, rules)
        if g != rule["color"]:
            key = (e["lv"], rgb, e.get("c"))
            bad[key].append(e["id"])
        if rule.get("dashed") and kind_of(e["t"]) == "line" and e.get("ls") in ("0", ""):
            solid_house.append(e["id"])
    for (lv, rgb, idx), ids in sorted(bad.items(), key=lambda kv: -len(kv[1])):
        exp = pl21[str(lv)]["color"]
        out.append(_f("symbology", "warning",
                      f"Level {lv} ({pl21[str(lv)]['name']}): {len(ids)} đối tượng màu "
                      f"{'#%02X%02X%02X' % rgb if rgb else '?'} (chỉ số {idx}), PL22 yêu cầu {groups[exp]['label']}",
                      lv, ids))
    if solid_house:
        out.append(_f("symbology", "warning",
                      f"{len(solid_house)} đối tượng tường nhà (level 14) vẽ nét liền; PL22 I.4b: ranh nhà vẽ nét đứt "
                      f"(trừ đoạn trùng ranh thửa thì vẽ ranh thửa nét liền ở level 10)", 14, solid_house))
    return out


# -------------------------------------------------------------- topology ---
def _rk(p, tol) -> Tuple[int, int]:
    return (round(p[0] / tol), round(p[1] / tol))


def build_lines(els, levels) -> List[Tuple[LineString, Dict[str, Any]]]:
    out = []
    for e in els:
        if e["lv"] in levels and "v" in e and len(e["v"]) >= 2 and kind_of(e["t"]) == "line":
            v = e["v"]
            if e["t"] in (6, 14) and v[0] != v[-1]:
                v = v + [v[0]]
            try:
                ls = LineString(v)
                if ls.length > 0:
                    out.append((ls, e))
            except Exception:  # noqa: BLE001
                pass
    return out


def check_topology(els, rules, boundary_levels) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """Trả về (findings, mạng đã nối nút để dựng thửa)."""
    topo = rules["topology"]
    tol = topo["snap_tolerance_m"]
    gap_tol = topo["gap_tolerance_m"]
    short = topo["short_segment_m"]
    out: List[Dict[str, Any]] = []
    lines = build_lines(els, boundary_levels)
    if not lines:
        return [_f("topology", "warning", f"Không có đường ranh nào trên các level {sorted(boundary_levels)}")], {}

    geoms = [g for g, _ in lines]
    tree = STRtree(geoms)

    # 1) đoạn quá ngắn, tự cắt, trùng đoạn
    seg_owner = defaultdict(set)
    vertex_keys = set()
    for g, e in lines:
        cs = list(g.coords)
        for p in cs:
            vertex_keys.add(_rk(p, tol))
        for a, b in zip(cs, cs[1:]):
            d = math.dist(a, b)
            if d < short and d > 0:
                out.append(_f("topology", "warning", f"Đoạn quá ngắn {d * 1000:.1f} mm", e["lv"], [e["id"]], a))
            k = tuple(sorted((_rk(a, tol), _rk(b, tol))))
            if k[0] != k[1]:
                seg_owner[k].add(e["id"])
        if not g.is_simple and not (g.is_ring and len(cs) > 3):
            out.append(_f("topology", "error", "Đường tự cắt chính nó", e["lv"], [e["id"]], cs[0]))
    dup = [ids for ids in seg_owner.values() if len(ids) > 1]
    dup_pairs = Counter(tuple(sorted(ids)) for ids in dup)
    for ids, n in dup_pairs.most_common():
        g0 = next(g for g, e in lines if e["id"] == ids[0])
        out.append(_f("topology", "error", f"Trùng đường: {n} đoạn trùng nhau giữa các đối tượng",
                      None, list(ids), g0.coords[0]))

    # 2) điểm treo / hở (đầu mút không nối với gì)
    #    Đầu mút nằm trên rìa khu đo (bao lồi của mạng ranh) hoặc chạm khung (level 63) -> chỉ là info.
    rim = unary_union(geoms).convex_hull.exterior
    frame = [g for g, _ in build_lines(els, {topo.get("frame_level", 63)})]
    frame_u = unary_union(frame) if frame else None
    end_count = Counter()
    for g, e in lines:
        if g.is_ring:
            continue
        end_count[_rk(g.coords[0], tol)] += 1
        end_count[_rk(g.coords[-1], tol)] += 1
    for i, (g, e) in enumerate(lines):
        if g.is_ring:
            continue
        for p in (g.coords[0], g.coords[-1]):
            if end_count[_rk(p, tol)] > 1:
                continue
            pt = Point(p)
            # nằm trên thân đường khác?
            best, best_id, best_j = None, None, None
            for j in tree.query(pt.buffer(gap_tol)):
                if j == i:
                    continue
                d = geoms[j].distance(pt)
                if d <= gap_tol and (best is None or d < best):
                    best, best_id, best_j = d, lines[j][1]["id"], j
            if best is None and (rim.distance(pt) < 1.0 or (frame_u is not None and frame_u.distance(pt) < gap_tol)):
                out.append(_f("topology", "info", "Đầu mút tự do ở rìa khu đo / chạm khung (thường là ranh khu đo)",
                              e["lv"], [e["id"]], p))
                continue
            if best is not None and best <= tol:
                if _rk(p, tol) not in _vertices_of(geoms[best_j], tol):
                    out.append(_f("topology", "error",
                                  "Điểm cuối chạm thân đường khác nhưng KHÔNG có điểm nút (16.1b)",
                                  e["lv"], [e["id"], best_id], p))
                continue
            if best is not None:
                out.append(_f("topology", "error", f"Hở {best * 100:.1f} cm - chưa bắt điểm vào đường gần nhất",
                              e["lv"], [e["id"], best_id], p, gap_m=round(best, 4)))
            else:
                out.append(_f("topology", "error", "Điểm treo (đầu mút không nối với đường nào)",
                              e["lv"], [e["id"]], p))

    # 3) giao cắt không có điểm nút
    noded = unary_union(geoms)
    segs = list(noded.geoms) if isinstance(noded, MultiLineString) else [noded]
    node_deg = Counter()
    for s in segs:
        node_deg[_rk(s.coords[0], tol)] += 1
        node_deg[_rk(s.coords[-1], tol)] += 1
    missing = [k for k, d in node_deg.items() if d >= 3 and k not in vertex_keys]
    for k in missing[:500]:
        p = (k[0] * tol, k[1] * tol)
        ids = [lines[j][1]["id"] for j in tree.query(Point(p).buffer(tol * 2))]
        out.append(_f("topology", "error", "Hai đường cắt nhau nhưng không có điểm nút tại giao điểm (16.1b)",
                      None, ids, p))
    return out, {"segments": segs, "lines": lines}


def _vertices_of(g, tol):
    return {_rk(p, tol) for p in g.coords}


# --------------------------------------------------------------- parcels ---
_num = re.compile(r"^\s*(\d+)\s*$")


def _area_value(s: str) -> Optional[float]:
    s = s.strip().replace(" ", "")
    if not re.match(r"^\d+([.,]\d+)?$", s):
        return None
    return float(s.replace(",", "."))


def build_parcels(els, rules, net) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """Dựng thửa từ mạng ranh, gán nhãn, so diện tích. Trả về (parcels, findings)."""
    lab = rules["labels"]
    out: List[Dict[str, Any]] = []
    segs = net.get("segments") or []
    if not segs:
        return [], out
    polys, cuts, dangles, invalid = polygonize_full(segs)
    polys = [p for p in polys.geoms if p.area > 0.01]
    if not polys:
        return [], [_f("parcels", "error", "Không dựng được thửa nào từ ranh - mạng ranh chưa khép kín")]
    ptree = STRtree(polys)

    def texts(level):
        return [e for e in els if e["lv"] == level and "tx" in e and "o" in e]

    def assign(items):
        res = defaultdict(list)
        outside = []
        for e in items:
            pt = Point(e["o"])
            hit = [i for i in ptree.query(pt) if polys[i].covers(pt)]
            if hit:
                res[min(hit, key=lambda i: polys[i].area)].append(e)
            else:
                outside.append(e)
        return res, outside

    nums, nums_out = assign([e for e in texts(lab["number_level"]) if _num.match(e["tx"])])
    lts, lts_out = assign(texts(lab["landtype_level"]))
    areas, areas_out = assign(texts(lab["area_level"]))
    valid_codes = set(rules.get("landtype_codes", []))

    parcels = []
    for i, p in enumerate(polys):
        c = p.representative_point()
        n = nums.get(i, [])
        lt = lts.get(i, [])
        ar = areas.get(i, [])
        rec = {"parcel": i + 1, "area_geom": round(p.area, 1), "xy": [round(c.x, 3), round(c.y, 3)],
               "so_thua": [e["tx"].strip() for e in n], "loai_dat": [e["tx"].strip() for e in lt],
               "dien_tich_nhan": [e["tx"].strip() for e in ar],
               "label_ids": [e["id"] for e in n + lt + ar]}
        parcels.append(rec)
        small = p.area < lab["small_parcel_m2"]
        if not n:
            out.append(_f("parcels", "warning" if small else "error",
                          f"Thửa {p.area:.1f} m² không có số thửa bên trong"
                          + (" (thửa nhỏ - kiểm tra bảng thửa nhỏ hẹp / nhãn đặt ngoài có mũi tên)" if small else ""),
                          None, None, (c.x, c.y)))
        elif len(n) > 1:
            out.append(_f("parcels", "error",
                          f"Một vùng có {len(n)} số thửa {rec['so_thua']} - thiếu ranh chia thửa hoặc nhãn đặt sai",
                          None, [e["id"] for e in n], (c.x, c.y)))
        if n and not lt:
            out.append(_f("parcels", "warning", f"Thửa {rec['so_thua'][0]} thiếu nhãn loại đất (level {lab['landtype_level']})",
                          None, [e["id"] for e in n], (c.x, c.y)))
        for e in lt:
            for code in re.split(r"[+\s]+", e["tx"].strip().upper()):
                if code and valid_codes and code not in valid_codes:
                    out.append(_f("parcels", "warning",
                                  f"Mã loại đất '{code}' không có trong danh mục cấu hình (đối chiếu TT 08/2024)",
                                  e["lv"], [e["id"]], e["o"]))
        if n and not ar and not small:
            out.append(_f("parcels", "warning", f"Thửa {rec['so_thua'][0]} thiếu nhãn diện tích (level {lab['area_level']})",
                          None, [e["id"] for e in n], (c.x, c.y)))
        for e in ar:
            v = _area_value(e["tx"])
            if v is None:
                out.append(_f("parcels", "warning", f"Nhãn diện tích không đúng định dạng: '{e['tx']}'",
                              e["lv"], [e["id"]], e["o"]))
                continue
            diff = v - p.area
            rec["chenh_lech_m2"] = round(diff, 2)
            if abs(diff) > lab["area_tolerance_m2"]:
                expect = f"{round(p.area, 1):.1f}".replace(".", ",")
                sev = "warning" if abs(diff) < 1.0 else "error"
                why = "làm tròn chưa đúng 0,1 m² (13.5b)" if abs(diff) < 1.0 else "nhãn không khớp hình thửa"
                out.append(_f("parcels", sev,
                              f"Thửa {rec['so_thua'][0] if n else '?'}: nhãn '{e['tx'].strip()}' nhưng diện tích tính "
                              f"được {p.area:.2f} m² -> nhãn đúng phải là '{expect}' ({why})",
                              None, [e["id"]], (c.x, c.y), area_label=v, area_geom=round(p.area, 2),
                              area_expected=expect))

    for e in nums_out:
        out.append(_f("parcels", "warning", f"Số thửa '{e['tx']}' nằm ngoài mọi thửa dựng được", e["lv"], [e["id"]], e["o"]))

    # trùng số thửa, thiếu số trong dãy
    allnums = Counter(int(_num.match(e["tx"]).group(1)) for lst in nums.values() for e in lst)
    allnums.update(int(_num.match(e["tx"]).group(1)) for e in nums_out)
    for k, cnt in sorted(allnums.items()):
        if cnt > 1:
            ids = [e["id"] for lst in list(nums.values()) + [nums_out] for e in lst
                   if _num.match(e["tx"]).group(1) == str(k)]
            out.append(_f("parcels", "error", f"Số thửa {k} bị trùng {cnt} lần trong tờ", None, ids))
    if allnums:
        mx = max(allnums)
        gaps = [k for k in range(1, mx + 1) if k not in allnums]
        if gaps:
            out.append(_f("parcels", "warning",
                          f"Dãy số thửa 1..{mx} thiếu {len(gaps)} số: {gaps[:40]}{' ...' if len(gaps) > 40 else ''}"))
    return parcels, out


# ------------------------------------------------------------------ run ---
def run_checks(data, rules, checks=None) -> Dict[str, Any]:
    t0 = time.time()
    els = [e for e in data["elements"] if e.get("r", [0])[0] > -1e6]  # bỏ phần tử điều khiển pattern
    levels_by_code = {l["code"]: l for l in data["header"].get("levels", [])}
    checks = list(checks or ALL_CHECKS)
    findings: List[Dict[str, Any]] = []
    parcels: List[Dict[str, Any]] = []
    if "levels" in checks:
        findings += check_levels(els, rules)
    if "symbology" in checks:
        findings += check_symbology(els, rules, levels_by_code)
    net = {}
    if "topology" in checks or "parcels" in checks:
        b = {int(k) for k, v in rules["levels"].items() if v.get("boundary")}
        if rules["topology"].get("include_admin_boundary"):
            b |= set(rules["topology"].get("admin_levels", []))
        topo_f, net = check_topology(els, rules, b)
        if "topology" in checks:
            findings += topo_f
    if "parcels" in checks:
        parcels, pf = build_parcels(els, rules, net)
        findings += pf
    sev = Counter(f["severity"] for f in findings)
    by_check = Counter(f["check"] for f in findings)
    return {"findings": findings, "parcels": parcels,
            "stats": {"errors": sev.get("error", 0), "warnings": sev.get("warning", 0), "info": sev.get("info", 0),
                      "by_check": dict(by_check), "parcels_built": len(parcels),
                      "seconds": round(time.time() - t0, 2)}}


def write_report(result: Dict[str, Any], dgn_path: str, out_dir: str) -> Dict[str, str]:
    os.makedirs(out_dir, exist_ok=True)
    base = os.path.splitext(os.path.basename(dgn_path or "drawing"))[0]
    ts = time.strftime("%Y%m%d_%H%M%S")
    f1 = os.path.join(out_dir, f"{base}_KCS_{ts}.csv")
    with open(f1, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["STT", "Hạng mục", "Mức", "Nội dung", "Level", "X", "Y", "ID đối tượng"])
        for i, x in enumerate(result["findings"], 1):
            xy = x.get("xy") or ["", ""]
            w.writerow([i, x["check"], x["severity"], x["message"], x.get("level", ""), xy[0], xy[1],
                        " ".join(x.get("ids", []))])
    f2 = os.path.join(out_dir, f"{base}_THUA_{ts}.csv")
    with open(f2, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Vùng", "Số thửa", "Loại đất", "DT nhãn", "DT hình học", "Chênh lệch", "X", "Y"])
        for p in result["parcels"]:
            w.writerow([p["parcel"], "/".join(p["so_thua"]), "/".join(p["loai_dat"]), "/".join(p["dien_tich_nhan"]),
                        p["area_geom"], p.get("chenh_lech_m2", ""), p["xy"][0], p["xy"][1]])
    return {"findings_csv": f1, "parcels_csv": f2}
