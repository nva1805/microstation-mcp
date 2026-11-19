"""
src/tools.py
Khai báo toàn bộ công cụ MCP.

Nguyên tắc:
- Đọc: mặc định đọc file đang mở; truyền `file` để mở ngầm CHỈ ĐỌC một file khác
  (không đổi file người dùng đang xem).
- Ghi: chỉ trên file đang mở; mặc định dry_run=True (xem trước). Trước lần ghi đầu tiên
  vào mỗi file, tự sao lưu; mỗi lô lệnh là một bước undo (ms_undo).
"""

from __future__ import annotations

import json
import os
import shutil
import time
from typing import Any, Dict, List, Literal, Optional

from fastmcp import FastMCP
from fastmcp.utilities.types import Image

from src.core import analysis, kcs
from src.core.com import com
from src.core.engine import WORK, engine, hexenc
from src.core import render, screenshot

_backed_up: Dict[str, str] = {}

KEYIN_DENY = ("vba", "mdl", "exit", "quit", "close", "compress", "newfile", "delete", "rd=", "save",
              "reference detach", "level delete", "dialog", "uninstall", "workmode")


def _num(v: Optional[float]) -> str:
    return "" if v is None else repr(float(v))


def _ms_windows() -> List[str]:
    """Tiêu đề cửa sổ chính của mọi phiên MicroStation đang chạy (lớp MstnTop)."""
    import win32gui
    out: List[str] = []
    win32gui.EnumWindows(lambda h, _: out.append(win32gui.GetWindowText(h))
                         if win32gui.GetClassName(h) == "MstnTop" and win32gui.IsWindowVisible(h) else None, None)
    return out


def _active_path() -> str:
    return com.run(lambda: com.app().ActiveDesignFile.FullName)


def _active_levels() -> Dict[int, str]:
    d = engine.dump(types=[-1], geometry=False)  # chỉ lấy header
    return {l["code"]: l["name"] for l in d["header"].get("levels", [])}


def _sym(c: Dict[str, Any]) -> List[str]:
    return ["" if c.get("level") is None else str(int(c["level"])),
            "" if c.get("color") is None else str(int(c["color"])),
            "" if c.get("weight") is None else str(int(c["weight"])),
            "" if c.get("style") is None else str(c["style"])]


def _pts(points: List[List[float]]) -> str:
    return ";".join(f"{float(p[0])!r},{float(p[1])!r}" for p in points)


_GROUP_INDEX = {"black": 0, "green": 2, "red": 3}  # chỉ số màu chuẩn MicroStation cho nhóm màu PL22


def _defaults(c: Dict[str, Any], rules: Dict[str, Any]) -> Dict[str, Any]:
    """Lệnh TẠO mới: nếu không chỉ định thì lấy màu theo nhóm PL22 của level, nét 0, kiểu nét 0
    (tránh thừa hưởng thuộc tính active đang đặt trong MicroStation)."""
    if c.get("op") not in ("line", "linestring", "shape", "text", "cell", "join"):
        return c
    c = dict(c)
    if c.get("level") is not None and c.get("color") is None:
        g = rules["levels"].get(str(int(c["level"])), {}).get("color")
        if g in _GROUP_INDEX:
            c["color"] = _GROUP_INDEX[g]
    c.setdefault("weight", 0)
    if c.get("op") != "cell":
        c.setdefault("style", "0")
    return c


def _to_tsv(c: Dict[str, Any]) -> str:
    op = c.get("op")
    if op in ("line", "linestring", "shape"):
        pts = c.get("points") or []
        need = 2 if op != "shape" else 3
        if len(pts) < need:
            raise ValueError(f"{op} cần ít nhất {need} điểm")
        if op == "shape" and pts[0] != pts[-1]:
            pts = pts + [pts[0]]
        return "\t".join([op] + _sym(c) + [_pts(pts)])
    if op == "text":
        if not c.get("text"):
            raise ValueError("text cần 'text'")
        return "\t".join(["text"] + _sym(c) + [f"{float(c['x'])!r},{float(c['y'])!r}", _num(c.get("height")),
                                               _num(c.get("rotation", 0)), hexenc(str(c["text"]))])
    if op == "cell":
        return "\t".join(["cell"] + _sym(c) + [f"{float(c['x'])!r},{float(c['y'])!r}", _num(c.get("scale", 1)),
                                               _num(c.get("rotation", 0)), hexenc(str(c["name"]))])
    if op == "set":
        return "\t".join(["set", str(c["id"])] + _sym(c))
    if op == "move":
        return "\t".join(["move", str(c["id"]), f"{float(c['dx'])!r},{float(c['dy'])!r}"])
    if op == "text_set":
        return "\t".join(["text_set", str(c["id"]), hexenc(str(c["text"]))])
    if op == "delete":
        return "\t".join(["delete", str(c["id"])])
    if op == "join":
        ids = c.get("ids") or []
        if len(ids) < 2:
            raise ValueError("join cần ít nhất 2 id")
        return "\t".join(["join"] + _sym(c) + [",".join(map(str, ids)), "shape" if c.get("as", "shape") == "shape" else "chain",
                                               _num(c.get("gap", 0.001)), "1" if c.get("delete_originals") else "0"])
    raise ValueError(f"op không hợp lệ: {op}")


def register(server: FastMCP) -> None:
    # ================================================================ status ===
    @server.tool()
    def ms_status() -> Dict[str, Any]:
        """Trạng thái kết nối: phiên bản MicroStation, file đang mở, bộ máy đọc/ghi (vba nhanh | com dự phòng)."""
        def _s():
            app = com.app(require_file=False)
            has = bool(app.HasActiveDesignFile)
            return {"microstation": str(app.Version), "has_file": has,
                    "active_file": app.ActiveDesignFile.FullName if has else None,
                    "active_model": app.ActiveModelReference.Name if has else None}
        st = com.run(_s)
        wins = _ms_windows()
        if len(wins) > 1:
            st["warning"] = (f"Đang chạy {len(wins)} phiên MicroStation: {wins}. MCP chỉ điều khiển phiên đăng ký COM "
                             "(thường là phiên mở đầu tiên) - nên đóng bớt để tránh ghi nhầm file.")
        if st["has_file"]:
            st["engine"] = engine.ensure()
            if engine.last_error:
                st["engine_error"] = engine.last_error
        st["work_dir"] = WORK
        return st

    # ================================================================== file ===
    @server.tool()
    def ms_open(path: str, read_only: bool = False) -> Dict[str, Any]:
        """Mở file .dgn trên CỬA SỔ MicroStation (thay file người dùng đang xem).
        Chỉ để đọc thì KHÔNG cần tool này - dùng tham số `file` của ms_summary/ms_query/kcs_check."""
        if not os.path.exists(path):
            raise ValueError(f"Không thấy file: {path}")
        p = os.path.abspath(path)
        name = os.path.basename(p).lower()
        if len(_ms_windows()) > 1 and any(t.lower().startswith(name + " [") for t in _ms_windows()):
            raise RuntimeError(
                f"'{os.path.basename(p)}' đang mở ở một phiên MicroStation khác - mở thêm sẽ bật hộp thoại "
                "'File is currently in use' và treo. Hãy đóng bớt một phiên MicroStation.")
        com.run(lambda: com.app(require_file=False).OpenDesignFile(p, read_only))
        engine.reset()
        return {"opened": p, "read_only": read_only}

    @server.tool()
    def ms_save() -> Dict[str, Any]:
        """Lưu file đang mở (V8i thường đã tự ghi thay đổi; dùng khi cần chắc chắn)."""
        com.run(lambda: com.app().ActiveDesignFile.Save())
        return {"saved": _active_path()}

    # ================================================================= read ====
    @server.tool()
    def ms_summary(file: Optional[str] = None) -> Dict[str, Any]:
        """Thống kê TOÀN BỘ bản vẽ theo level: số đối tượng theo loại, màu (chỉ số=RGB), lực nét, kiểu nét,
        mẫu chữ, cell, độ dài đường; đối chiếu tên lớp PL21 và đánh dấu level không có trong PL21.
        file=None: file đang mở. file=đường dẫn: mở ngầm chỉ đọc (không đổi file đang xem)."""
        rules = analysis.load_rules()
        data = engine.dump(file=file, geometry=True, style=2)
        s = analysis.summarize(data, rules)
        if file:
            s["note"] = ("File mở ngầm: MicroStation không cho đọc tên/thuộc tính level (giới hạn V8i), "
                         "lực nét/màu theo level (ByLevel, giá trị -1) không giải được.")
        return s

    @server.tool()
    def ms_query(file: Optional[str] = None, levels: Optional[List[int]] = None,
                 types: Optional[List[str]] = None, bbox: Optional[List[float]] = None,
                 text: Optional[str] = None, text_regex: Optional[str] = None, cell: Optional[str] = None,
                 color: Optional[int] = None, include_geometry: bool = True,
                 include_style: bool = False, offset: int = 0, limit: int = 100) -> Dict[str, Any]:
        """Tìm đối tượng theo bộ lọc, có phân trang.
        levels: [10, 13]; types: ['Line','LineString','Shape','Text','TextNode','Cell',...] hoặc 'text'/'cell'/'line';
        bbox: [xmin, ymin, xmax, ymax]; text: chứa chuỗi (không phân biệt hoa thường); text_regex: biểu thức chính quy;
        cell: tên cell; color: chỉ số màu; include_style: đọc tên kiểu nét (chậm hơn, ~0,5 ms/đối tượng
        với kiểu nét tùy biến - mặc định tắt, style='?'). Trả về id, loại, level, màu, nét, range, đỉnh/chữ/cell, độ dài, diện tích (nếu khép kín)."""
        data = engine.dump(file=file, levels=levels, types=analysis.parse_types(types), bbox=bbox,
                           geometry=include_geometry, style=1 if include_style else 0)
        els = analysis.filter_elements(data["elements"], text, text_regex, cell, color)
        page = els[offset: offset + max(1, min(limit, 1000))]
        return {"total": len(els), "offset": offset, "returned": len(page), "engine": data["engine"],
                "seconds": data["seconds"],
                "elements": [analysis.compact(e, include_geometry) for e in page]}

    @server.tool()
    def ms_get(ids: List[str], file: Optional[str] = None, max_vertices: int = 2000) -> Dict[str, Any]:
        """Chi tiết đầy đủ của các đối tượng theo id (toàn bộ đỉnh, chữ, cell, range...)."""
        data = engine.dump(file=file, ids=ids, geometry=True, style=1)
        return {"elements": [analysis.compact(e, True, vmax=max_vertices) for e in data["elements"]],
                "missing": data["missing"]}

    # ================================================================ write ====
    @server.tool()
    def ms_edit(commands: List[Dict[str, Any]], dry_run: bool = True) -> Dict[str, Any]:
        """Ghi vào FILE ĐANG MỞ. Mặc định dry_run=True: chỉ kiểm tra và xem trước, chưa ghi.
        Khi dry_run=False: tự sao lưu file (lần đầu), đặt điểm undo, thực hiện cả lô, trả id mới.
        Mỗi lệnh là một dict có 'op':
          line|linestring|shape: points=[[x,y],...], level, color, weight, style
          text: text, x, y, height, rotation(độ), level, color...
          cell: name, x, y, scale, rotation (cần thư viện cell đã attach)
          set: id, level?, color?, weight?, style?   (đổi lớp/màu/nét)
          move: id, dx, dy      text_set: id, text      delete: id
          join: ids=[...], as='shape'|'chain', gap=0.001, delete_originals=false, level...  (nối đường thành vùng/chuỗi)
        level là SỐ level (code). style là tên kiểu nét ('0'..'7' hoặc tên line style).
        Lệnh tạo mới không ghi color/weight/style thì dùng mặc định PL22: màu theo nhóm của level
        (đen 0, xanh 2, đỏ 3), weight 0, style '0' - không thừa hưởng thuộc tính active của MicroStation."""
        if not commands:
            raise ValueError("Không có lệnh nào")
        rules = analysis.load_rules()
        commands = [_defaults(c, rules) for c in commands]
        tsv = [_to_tsv(c) for c in commands]
        levels = _active_levels()
        bad_lv = sorted({int(c["level"]) for c in commands if c.get("level") is not None} - set(levels))
        if bad_lv:
            raise ValueError(f"Level không tồn tại trong file: {bad_lv}")
        target = _active_path()
        ids = [str(c["id"]) for c in commands if "id" in c] + [str(i) for c in commands for i in c.get("ids", [])]
        preview: Dict[str, Any] = {"file": target, "n_commands": len(commands),
                                   "ops": dict(__import__("collections").Counter(c["op"] for c in commands))}
        if ids:
            cur = engine.dump(ids=ids, geometry=False)
            preview["affected_now"] = [analysis.compact(e, False) for e in cur["elements"][:50]]
            if cur["missing"]:
                raise ValueError(f"Không tìm thấy id: {cur['missing']}")
        if dry_run:
            preview["dry_run"] = True
            preview["next"] = "Gọi lại với dry_run=false để ghi thật."
            return preview
        if target not in _backed_up:
            bdir = os.path.join(WORK, "backups")
            os.makedirs(bdir, exist_ok=True)
            dst = os.path.join(bdir, f"{os.path.splitext(os.path.basename(target))[0]}_{time.strftime('%Y%m%d_%H%M%S')}.dgn")
            shutil.copy2(target, dst)
            _backed_up[target] = dst
        res = engine.apply(tsv)
        ok = [r for r in res if r.get("ok")]
        return {"file": target, "backup": _backed_up[target], "applied": len(ok), "failed": len(res) - len(ok),
                "results": [{"i": r["i"], "op": commands[r["i"] - 1]["op"], "ok": r["ok"],
                             **({"id": r["id"]} if "id" in r else {}), **({"error": r["err"]} if "err" in r else {})}
                            for r in res],
                "undo": "ms_undo hoàn tác toàn bộ lô này"}

    @server.tool()
    def ms_undo() -> Dict[str, Any]:
        """Hoàn tác về điểm undo đặt trước lô ms_edit gần nhất (key-in 'undo mark')."""
        com.run(lambda: com.app().CadInputQueue.SendKeyin("undo mark"))
        return {"undone": True}

    # ================================================================== KCS ====
    @server.tool()
    def kcs_check(file: Optional[str] = None, checks: Optional[List[str]] = None,
                  max_findings: int = 150, write_csv: bool = True,
                  severity: Optional[List[str]] = None) -> Dict[str, Any]:
        """Kiểm tra KCS theo TT 26/2024 (PL21 phân lớp, PL22 màu/nét, 16.1b topology, Điều 13 thửa/nhãn/diện tích).
        checks: tập con của ['levels','symbology','topology','parcels'] (mặc định tất cả).
        - levels: level ngoài PL21, level đã bãi bỏ, dạng đối tượng sai lớp
        - symbology: màu sai nhóm (đen/xanh/nâu/đỏ), tường nhà nét liền
        - topology: điểm treo, hở, trùng đường, cắt nhau không có nút, tự cắt, đoạn quá ngắn
        - parcels: dựng thửa từ ranh (10,21,23,31,32,38 + địa giới), thiếu/trùng số thửa, thiếu nhãn,
          mã loại đất lạ, diện tích nhãn lệch hình học, dãy số thửa thiếu
        Ghi đầy đủ ra CSV (mở bằng Excel) và trả về tối đa max_findings dòng."""
        rules = analysis.load_rules()
        dashed = [int(k) for k, v in rules["levels"].items() if v.get("dashed")]
        want_style = not checks or "symbology" in checks
        data = engine.dump(file=file, geometry=True, style=0, style_levels=dashed if want_style else None)
        r = kcs.run_checks(data, rules, checks)
        f = r["findings"]
        if severity:
            f = [x for x in f if x["severity"] in severity]
        order = {"error": 0, "warning": 1, "info": 2}
        f.sort(key=lambda x: (order.get(x["severity"], 3), x["check"]))
        out = {"file": data["header"].get("file"), "stats": r["stats"], "findings": f[:max_findings],
               "findings_total": len(f), "read_seconds": data["seconds"]}
        if write_csv:
            out["report"] = kcs.write_report(r, data["header"].get("file", ""), os.path.join(WORK, "reports"))
        return out

    # ================================================================= view ====
    def _render_img(file, bbox, levels=None, highlight=(), width=1400, title=None):
        d = engine.dump(file=file, levels=levels, bbox=bbox, geometry=True, style=2)
        png = render.render(d["elements"], bbox, width=width, highlight=highlight, fmt="png", title=title)
        return Image(data=png, format="png")

    @server.tool()
    def ms_view(action: Literal["fit", "zoom_window", "zoom_to", "zoom_ids"], view: int = 1,
                bbox: Optional[List[float]] = None, x: Optional[float] = None, y: Optional[float] = None,
                radius: float = 20.0, ids: Optional[List[str]] = None, shot: bool = True,
                source: Literal["render", "window"] = "render") -> Any:
        """Đưa khung nhìn MicroStation tới vùng cần xem và trả kèm ảnh vùng đó.
        fit: toàn bộ; zoom_window: bbox=[xmin,ymin,xmax,ymax]; zoom_to: tâm x,y bán kính radius (m);
        zoom_ids: phóng tới các đối tượng theo id (tô đậm chúng trong ảnh) - dùng với kcs_check.
        source='render' (mặc định): ảnh DỰNG TỪ DỮ LIỆU file đang mở - luôn đúng vùng kể cả khi cửa sổ
        MicroStation bị che/thu nhỏ (không gồm file tham chiếu). source='window': chụp cửa sổ thật
        (chỉ đúng khi cửa sổ MicroStation đang hiện trên màn hình)."""
        if action == "zoom_ids":
            d = engine.dump(ids=ids or [], geometry=False)
            rs = [e["r"] for e in d["elements"]]
            if not rs:
                raise ValueError("Không tìm thấy đối tượng")
            bbox = [min(r[0] for r in rs) - radius, min(r[1] for r in rs) - radius,
                    max(r[2] for r in rs) + radius, max(r[3] for r in rs) + radius]
        elif action == "zoom_to":
            bbox = [x - radius, y - radius, x + radius, y + radius]

        def _do():
            app = com.app()
            v = app.ActiveDesignFile.Views(view)
            if action == "fit":
                v.Fit(False)
            else:
                ext = v.Extents
                aspect = ext.X / ext.Y if ext.Y else 1.0
                w, h = max(bbox[2] - bbox[0], 0.5), max(bbox[3] - bbox[1], 0.5)
                if w / h > aspect:
                    h = w / aspect
                else:
                    w = h * aspect
                cx, cy = (bbox[0] + bbox[2]) / 2, (bbox[1] + bbox[3]) / 2
                v.Origin = app.Point3dFromXYZ(cx - w / 2, cy - h / 2, v.Origin.Z)
                v.Extents = app.Point3dFromXYZ(w, h, ext.Z)
            v.Redraw()
            o, e = v.Origin, v.Extents
            return [o.X, o.Y, o.X + e.X, o.Y + e.Y]
        view_box = com.run(_do)
        if not shot:
            return {"ok": True, "view_bbox": view_box}
        if source == "window":
            time.sleep(0.3)
            data, what = screenshot.capture(view=view, fmt="jpeg")
            return Image(data=data, format="jpeg") if data else {"ok": True, "screenshot_error": what}
        return _render_img(None, view_box, highlight=ids or (), title=os.path.basename(_active_path()))

    @server.tool()
    def ms_render(file: Optional[str] = None, bbox: Optional[List[float]] = None,
                  levels: Optional[List[int]] = None, highlight: Optional[List[str]] = None,
                  width: int = 1400) -> Any:
        """Dựng ảnh bản vẽ từ dữ liệu mà KHÔNG đổi khung nhìn MicroStation. Dùng được cho file mở ngầm (file=).
        bbox=None: toàn bộ; levels: chỉ vẽ các level này; highlight: id cần tô đậm (màu hồng)."""
        return _render_img(file, bbox, levels, highlight or (), width,
                           title=os.path.basename(file) if file else None)

    @server.tool()
    def ms_screenshot(view: int = 1, full_window: bool = False,
                      format: Literal["png", "jpeg"] = "png", max_width: int = 1600) -> Any:
        """Chụp cửa sổ MicroStation thật (khung bản vẽ VpWindow). Chỉ đúng khi cửa sổ MicroStation đang hiển thị
        trên màn hình - nếu bị cửa sổ khác che, MicroStation không vẽ lại và ảnh có thể là ảnh cũ.
        Muốn ảnh luôn đúng: dùng ms_render / ms_view."""
        data, what = screenshot.capture(view=view, full_window=full_window, fmt=format, max_width=max_width)
        if not data:
            raise RuntimeError(what)
        return Image(data=data, format=format)

    # ================================================================= copy ====
    @server.tool()
    def ms_copy_from(file: str, levels: Optional[List[int]] = None, types: Optional[List[str]] = None,
                     bbox: Optional[List[float]] = None, ids: Optional[List[str]] = None,
                     skip_existing: bool = True, dry_run: bool = True) -> Dict[str, Any]:
        """Chép đối tượng từ một file .dgn khác (mở ngầm, chỉ đọc) vào FILE ĐANG MỞ - xử lý hoàn toàn trong server,
        dữ liệu không đi qua hội thoại. Lọc theo levels/types/bbox/ids.
        Giữ level, kiểu nét, lực nét. Màu: giữ chỉ số nếu là màu bảng (0-255); màu riêng (true color) của file
        nguồn quy về chỉ số cùng RGB trong file đích, không có thì dùng màu chuẩn PL22 của level.
        Hỗ trợ Line, LineString, Shape, ComplexChain/Shape (thành LineString/Shape), Text; loại khác được báo bỏ qua.
        skip_existing: bỏ đối tượng trùng hình học đã có ở file đích (cùng level).
        Mặc định dry_run=True (chỉ thống kê); dry_run=False để ghi (có sao lưu, ms_undo hoàn tác cả lô)."""
        if not os.path.exists(file):
            raise ValueError(f"Không thấy file: {file}")
        if os.path.normcase(os.path.abspath(file)) == os.path.normcase(_active_path()):
            raise ValueError("File nguồn trùng file đang mở (file đích). Hãy mở file cần ghi trên MicroStation.")
        rules = analysis.load_rules()
        src = engine.dump(file=file, levels=levels, types=analysis.parse_types(types), bbox=bbox, ids=ids,
                          geometry=True, style=1)
        dst = engine.dump(levels=levels, bbox=bbox, geometry=True)
        rgb_to_idx: Dict[int, int] = {}
        for i, v in enumerate(dst.get("palette") or []):
            rgb_to_idx.setdefault(v, i)

        def key(lv, pts):
            r = [(round(p[0], 3), round(p[1], 3)) for p in pts]
            return (lv, tuple(min(r, r[::-1])))

        have = {key(e["lv"], e["v"]) for e in dst["elements"] if "v" in e}
        cmds: List[Dict[str, Any]] = []
        skipped: Dict[str, int] = {}
        dup = 0
        for e in src["elements"]:
            t, lv = e["t"], e["lv"]
            c: Dict[str, Any] = {"level": lv}
            col = e.get("c", -1)
            if 0 <= col <= 255:
                c["color"] = col
            elif e.get("rgb", -1) in rgb_to_idx:
                c["color"] = rgb_to_idx[e["rgb"]]
            if (e.get("w") if e.get("w") is not None else -1) >= 0:
                c["weight"] = e["w"]
            ls = str(e.get("ls", ""))
            if ls not in ("", "?", "ByLevel"):
                c["style"] = ls
            if t in (3, 4, 12) and len(e.get("v", [])) >= 2:
                c.update(op="line" if len(e["v"]) == 2 else "linestring", points=e["v"])
            elif t in (6, 14) and len(e.get("v", [])) >= 3:
                c.update(op="shape", points=e["v"])
            elif t == 17 and e.get("tx") and "o" in e:
                c.update(op="text", text=e["tx"], x=e["o"][0], y=e["o"][1], height=e.get("h") or None)
                c.pop("style", None)
            else:
                k = analysis.type_name(t)
                skipped[k] = skipped.get(k, 0) + 1
                continue
            if skip_existing and "points" in c and key(lv, c["points"]) in have:
                dup += 1
                continue
            cmds.append(_defaults(c, rules))
        by: Dict[str, int] = {}
        for c in cmds:
            k = f'level {c["level"]} {c["op"]}'
            by[k] = by.get(k, 0) + 1
        out: Dict[str, Any] = {"source": src["header"].get("file"), "target": _active_path(),
                               "read_source_s": src["seconds"], "to_copy": len(cmds), "by_level_op": by,
                               "skipped_unsupported": skipped, "skipped_already_present": dup}
        if dry_run or not cmds:
            out["dry_run"] = dry_run
            return out
        bad_lv = sorted({c["level"] for c in cmds} - set(_active_levels()))
        if bad_lv:
            raise ValueError(f"File đích không có level: {bad_lv}")
        target = out["target"]
        if target not in _backed_up:
            bdir = os.path.join(WORK, "backups")
            os.makedirs(bdir, exist_ok=True)
            bpath = os.path.join(bdir, f"{os.path.splitext(os.path.basename(target))[0]}_{time.strftime('%Y%m%d_%H%M%S')}.dgn")
            shutil.copy2(target, bpath)
            _backed_up[target] = bpath
        t0 = time.time()
        res = engine.apply([_to_tsv(c) for c in cmds])
        warns: Dict[str, int] = {}
        for r in res:
            if r.get("warn"):
                warns[r["warn"]] = warns.get(r["warn"], 0) + 1
        out.update(applied=sum(1 for r in res if r.get("ok")),
                   failed=[{"i": r["i"], "error": r.get("err")} for r in res if not r.get("ok")][:20],
                   warnings=warns, new_ids_first=[r["id"] for r in res if r.get("id")][:20],
                   new_id_range=[next((r["id"] for r in res if r.get("id")), None), next((r["id"] for r in reversed(res) if r.get("id")), None)],
                   write_s=round(time.time() - t0, 2), backup=_backed_up[target],
                   undo="ms_undo hoàn tác toàn bộ lô này")
        return out

    @server.tool()
    def ms_keyin(command: str) -> Dict[str, Any]:
        """Gửi một key-in MicroStation (vd 'fit view extended', 'level set active 10', 'place line').
        Chặn các lệnh nguy hiểm (vba, mdl, delete, save, exit, close, compress...)."""
        low = command.strip().lower()
        if any(low.startswith(d) or f" {d}" in low for d in KEYIN_DENY):
            raise ValueError(f"Key-in bị chặn vì an toàn: {command}")
        com.run(lambda: com.app().CadInputQueue.SendKeyin(command))
        return {"sent": command}
