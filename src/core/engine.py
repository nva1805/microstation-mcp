"""
src/core/engine.py
Bộ máy đọc/ghi nhanh: chạy mã VBA (vba/McpFast.bas) BÊN TRONG tiến trình MicroStation.

Luồng: Python ghi req.txt -> key-in `vba run [msmcp]McpFast.Main` -> VBA ghi resp.txt
(JSON lines) rồi done.txt (token). Nếu VBA không dùng được thì rơi về COM (chậm hơn
nhiều nhưng cùng định dạng dữ liệu).

Mã VBA là file chữ trong repo; lần đầu (hoặc khi file .bas đổi) sẽ được dựng thành
project %LOCALAPPDATA%\\msmcp\\msmcp.mvba bằng các key-in chuẩn của MicroStation.
"""

from __future__ import annotations

import hashlib
import json
import logging
import os
import shutil
import time
import uuid
from typing import Any, Dict, Iterable, List, Optional

from src.core.com import com

logger = logging.getLogger("msmcp.engine")

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
BAS_SRC = os.path.join(REPO_ROOT, "vba", "McpFast.bas")
WORK = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "msmcp")
BAS_COPY = os.path.join(WORK, "McpFast.bas")
REQ, RESP, DONE = (os.path.join(WORK, n) for n in ("req.txt", "resp.txt", "done.txt"))

ELEMENT_TYPES = {
    2: "Cell", 3: "Line", 4: "LineString", 6: "Shape", 7: "TextNode", 11: "Curve",
    12: "ComplexChain", 14: "ComplexShape", 15: "Ellipse", 16: "Arc", 17: "Text",
    22: "PointString", 27: "BSplineCurve", 33: "Dimension", 35: "SharedCell", 36: "Multiline",
    37: "Tag",
}
TYPE_CODES = {v.lower(): k for k, v in ELEMENT_TYPES.items()}


def hexenc(s: str) -> str:
    """Mã hóa chuỗi thành các code unit UTF-16 dạng hex 4 ký tự (VBA giải mã bằng ChrW)."""
    b = s.encode("utf-16-le")
    return "".join(f"{b[i + 1]:02X}{b[i]:02X}" for i in range(0, len(b), 2))


def rgb_tuple(v: int) -> Optional[tuple]:
    """MicroStation trả RGB dạng 0xRRGGBB (đã kiểm chứng: màu 3 -> 0xFF0000 = đỏ)."""
    if v is None or v < 0:
        return None
    return ((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF)


class EngineError(RuntimeError):
    pass


class Engine:
    def __init__(self) -> None:
        self.mode: Optional[str] = None  # "vba" | "com"
        self.last_error: Optional[str] = None
        self._loaded_in_session = False
        self.project = "msmcp"

    # ------------------------------------------------------------- cài đặt ---
    def _sha(self) -> str:
        with open(BAS_SRC, "rb") as f:
            return hashlib.sha1(f.read()).hexdigest()

    def _keyin(self, app, cmd: str) -> None:
        app.CadInputQueue.SendKeyin(cmd)

    def ensure(self) -> str:
        """Đảm bảo bộ máy VBA sẵn sàng; trả về 'vba' hoặc 'com'.

        Project được đặt tên theo hash mã nguồn (msmcp_<sha8>): đổi mã -> project mới,
        không bao giờ phải 'vba unload' project cũ (có thể bật hộp thoại hỏi lưu và chặn MicroStation).
        """
        if self.mode == "vba" and self._loaded_in_session:
            return "vba"
        if os.environ.get("MSMCP_DISABLE_VBA") == "1":
            self.mode = "com"
            return "com"
        os.makedirs(WORK, exist_ok=True)
        self.project = f"msmcp_{self._sha()[:8]}"
        mvba = os.path.join(WORK, self.project + ".mvba")
        need_build = not os.path.exists(mvba)

        def _install():
            app = com.app()
            if need_build:
                shutil.copyfile(BAS_SRC, BAS_COPY)
                self._keyin(app, f"vba create {mvba}")
                self._keyin(app, f'vba execute VBE.VBProjects("{self.project}").VBComponents.Import "{BAS_COPY}"')
                self._keyin(app, f"vba save {self.project}")
            else:
                self._keyin(app, f"vba load {mvba}")

        try:
            com.run(_install)
            self.mode = "vba"
            res = self._call_vba({"op": "ping"}, timeout=20)
            if not any(r.get("kind") == "pong" for r in res):
                raise EngineError("VBA không phản hồi ping")
            for f in os.listdir(WORK):  # dọn project cũ không còn dùng (bỏ qua nếu đang được nạp)
                if f.startswith("msmcp") and f.endswith(".mvba") and f != self.project + ".mvba":
                    try:
                        os.remove(os.path.join(WORK, f))
                    except OSError:
                        pass
            self._loaded_in_session = True
            self.last_error = None
        except Exception as ex:  # noqa: BLE001
            logger.warning("Không dùng được bộ máy VBA, chuyển sang COM: %s", ex)
            self.last_error = str(ex)
            self.mode = "com"
        return self.mode

    def _recover_break(self) -> bool:
        """Nếu trình VBA đang dừng ở [break] (lỗi lúc chạy/biên dịch), gửi lệnh Run>Reset (control 228)."""
        try:
            import win32gui
            hits = []
            win32gui.EnumWindows(lambda h, _: hits.append(h) if "[break]" in win32gui.GetWindowText(h)
                                 and "Visual Basic" in win32gui.GetWindowText(h) else None, None)
            if not hits:
                return False
            com.run(lambda: self._keyin(com.app(require_file=False),
                                        "vba execute VBE.CommandBars.FindControl(, 228).Execute"), timeout=20)
            return True
        except Exception:  # noqa: BLE001
            return False

    def reset(self) -> None:
        self._loaded_in_session = False
        self.mode = None

    # ------------------------------------------------------------- gọi VBA ---
    def _call_vba(self, params: Dict[str, Any], commands: Optional[List[str]] = None,
                  timeout: float = 300) -> List[Dict[str, Any]]:
        tok = uuid.uuid4().hex
        lines = [f"tok={tok}"] + [f"{k}={v}" for k, v in params.items() if v is not None]
        if commands:
            lines.append("---")
            lines.extend(commands)
        for p in (RESP, DONE):
            try:
                os.remove(p)
            except FileNotFoundError:
                pass
        with open(REQ, "w", encoding="ascii", newline="\r\n") as f:
            f.write("\n".join(lines) + "\n")

        com.run(lambda: self._keyin(com.app(), f"vba run [{self.project}]McpFast.Main"))

        deadline = time.time() + timeout
        delay = 0.01
        while time.time() < deadline:
            try:
                with open(DONE, "r", encoding="ascii", errors="ignore") as f:
                    if f.read().strip() == tok:
                        break
            except FileNotFoundError:
                pass
            time.sleep(delay)
            delay = min(delay * 1.5, 0.2)
        else:
            recovered = self._recover_break()
            raise EngineError(
                "VBA không trả lời trong thời gian chờ"
                + (" - đã tự Reset trình gỡ lỗi VBA (lỗi trong mã VBA)." if recovered else
                   ". MicroStation có thể đang bận hoặc đang hiện hộp thoại - kiểm tra cửa sổ MicroStation.")
            )
        out: List[Dict[str, Any]] = []
        with open(RESP, "r", encoding="ascii", errors="replace") as f:
            for ln in f:
                ln = ln.strip()
                if ln:
                    out.append(json.loads(ln))
        errs = [r for r in out if r.get("kind") == "error"]
        if errs:
            raise EngineError(errs[0].get("err", "lỗi VBA"))
        return out

    # ------------------------------------------------------------- dump ------
    def dump(self, file: Optional[str] = None, levels: Optional[Iterable[int]] = None,
             types: Optional[Iterable[int]] = None, bbox: Optional[List[float]] = None,
             ids: Optional[Iterable[str]] = None, geometry: bool = True,
             vmax: int = 0, style: int = 0,
             style_levels: Optional[Iterable[int]] = None) -> Dict[str, Any]:
        """Đọc đối tượng. file=None -> file đang mở; ngược lại mở ngầm CHỈ ĐỌC.

        style: 0 không đọc kiểu nét (nhanh), 1 đọc tất cả, 2 lấy mẫu 30 đối tượng/level.
        style_levels: luôn đọc kiểu nét trên các level này. Kiểu nét không đọc được ghi là '?'.
        (Đọc kiểu nét tùy biến tốn ~0,5 ms/đối tượng - 21.000 đối tượng = ~10 giây.)"""
        mode = self.ensure()
        t0 = time.time()
        if mode == "vba":
            params = {
                "op": "dump",
                "file": hexenc(os.path.abspath(file)) if file else None,
                "levels": ",".join(str(int(x)) for x in levels) if levels else None,
                "types": ",".join(str(int(x)) for x in types) if types else None,
                "bbox": ",".join(repr(float(x)) for x in bbox) if bbox else None,
                "ids": ",".join(str(x) for x in ids) if ids else None,
                "geom": "1" if geometry else "0",
                "vmax": str(int(vmax)),
                "style": str(int(style)),
                "stylelv": ",".join(str(int(x)) for x in style_levels) if style_levels else None,
            }
            rows = self._call_vba(params)
        else:
            rows = com.run(lambda: _com_dump(file, levels, types, bbox, ids, geometry, vmax))
        data: Dict[str, Any] = {"header": {}, "palette": [], "elements": [], "missing": [],
                                "engine": mode}
        for r in rows:
            k = r.get("kind")
            if k == "el":
                data["elements"].append(r)
            elif k == "header":
                data["header"] = r
            elif k == "palette":
                data["palette"] = r.get("rgb", [])
            elif k == "missing":
                data["missing"].append(r["id"])
            elif k == "end":
                data["engine_sec"] = r.get("sec")
        data["seconds"] = round(time.time() - t0, 3)
        return data

    # ------------------------------------------------------------- apply -----
    def copy(self, file: str, jobs: List[str]) -> List[Dict[str, Any]]:
        """Chép sâu phần tử từ file (mở ngầm chỉ đọc) vào model đang mở bằng CopyElement."""
        if self.ensure() != "vba":
            raise EngineError(f"Chép cần bộ máy VBA. Lỗi khi nạp VBA: {self.last_error}")
        com.run(lambda: self._keyin(com.app(), "mark"))
        rows = self._call_vba({"op": "copy", "file": hexenc(os.path.abspath(file))}, commands=jobs)
        return [r for r in rows if r.get("kind") == "res"]

    def apply(self, commands: List[str]) -> List[Dict[str, Any]]:
        mode = self.ensure()
        if mode != "vba":
            raise EngineError(
                "Ghi dữ liệu cần bộ máy VBA (an toàn và có undo theo lô). "
                f"Lỗi khi nạp VBA: {self.last_error}")
        com.run(lambda: self._keyin(com.app(), "mark"))  # điểm undo cho cả lô
        rows = self._call_vba({"op": "apply"}, commands=commands)
        return [r for r in rows if r.get("kind") == "res"]


# ----------------------------------------------------------- COM fallback ---
def _com_dump(file, levels, types, bbox, ids, geometry, vmax) -> List[Dict[str, Any]]:
    """Cùng định dạng với VBA nhưng qua COM (chậm: ~vài ms mỗi thuộc tính)."""
    app = com.app(require_file=file is None)
    ext = bool(file)
    df = app.OpenDesignFileForProgram(os.path.abspath(file), True) if ext else app.ActiveDesignFile
    m = df.DefaultModelReference if ext else app.ActiveModelReference
    rows: List[Dict[str, Any]] = []

    def rgb(c):
        try:
            return int(m.InternalColorToRGBColor(c))
        except Exception:  # noqa: BLE001
            return -1

    lvls = []
    for i in range(df.Levels.Count):
        lv = df.Levels.Item(i + 1)
        try:
            lvls.append({"code": lv.Code, "name": lv.Name, "color": lv.ElementColor,
                         "rgb": rgb(lv.ElementColor), "weight": lv.ElementLineWeight, "style": ""})
        except Exception:  # noqa: BLE001
            pass
    rows.append({"kind": "header", "file": df.FullName, "model": m.Name, "external": ext,
                 "levels": lvls})
    rows.append({"kind": "palette", "rgb": [rgb(i) for i in range(256)]})
    lset = set(int(x) for x in levels) if levels else None
    tset = set(int(x) for x in types) if types else None

    def one(e):
        t = e.Type
        try:
            lc = e.Level.Code
        except Exception:  # noqa: BLE001
            lc = -1
        if tset and t not in tset:
            return
        if lset and lc not in lset:
            return
        r = e.Range
        box = [r.Low.X, r.Low.Y, r.High.X, r.High.Y]
        if bbox and (box[2] < bbox[0] or box[0] > bbox[2] or box[3] < bbox[1] or box[1] > bbox[3]):
            return
        d = {"kind": "el", "id": str(app.DLongToString(e.ID)), "t": t, "lv": lc,
             "c": e.Color, "rgb": rgb(e.Color), "w": e.LineWeight, "ls": "", "gg": 0, "r": box}
        try:
            if t in (3, 4, 6, 11) and geometry:
                vs = e.AsVertexList.GetVertices()
                d["v"] = [[p.X, p.Y] for p in vs][: vmax or None]
            elif t == 17:
                te = e.AsTextElement
                d.update(tx=te.Text, o=[te.Origin.X, te.Origin.Y], h=te.TextStyle.Height)
            elif t == 2:
                ce = e.AsCellElement
                d.update(cell=ce.Name, o=[ce.Origin.X, ce.Origin.Y])
        except Exception as ex:  # noqa: BLE001
            d["aerr"] = str(ex)
        rows.append(d)

    if ids:
        for i in ids:
            try:
                one(m.GetElementByID(app.DLongFromString(str(i))))
            except Exception:  # noqa: BLE001
                rows.append({"kind": "missing", "id": str(i)})
    else:
        sc = app.CreateObjectInMicroStation("MicroStationDGN.ElementScanCriteria")
        ee = m.Scan(sc)
        while ee.MoveNext():
            e = ee.Current
            if e.IsGraphical:
                one(e)
    if ext:
        df.Close()
    return rows


engine = Engine()
