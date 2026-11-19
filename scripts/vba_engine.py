"""
scripts/vba_engine.py
Dựng và KIỂM CHỨNG bộ máy VBA (kết nối nhanh) mà MCP nạp vào MicroStation.

    python scripts/vba_engine.py info     # in mã băm, tên project, đường dẫn, các key-in sẽ chạy
    python scripts/vba_engine.py build    # dựng/nạp project từ vba/McpFast.bas (cần MicroStation đang mở file)
    python scripts/vba_engine.py verify   # xuất module đang nạp trong MicroStation và so với vba/McpFast.bas

`verify` chứng minh mã đang chạy bên trong MicroStation trùng khớp từng dòng với file nguồn
công khai trong repo (không có gì khác được nạp).
"""

from __future__ import annotations

import difflib
import hashlib
import os
import sys
import tempfile
import time

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from src.core.com import com  # noqa: E402
from src.core.engine import BAS_COPY, BAS_SRC, WORK, engine  # noqa: E402


def _canon(line: str) -> str:
    """VBA không phân biệt hoa/thường và TỰ đổi hoa/thường tên định danh trong toàn project
    (vd idS -> ids, .X -> .x). Vì vậy: phần mã -> chữ thường; chuỗi "..." và chú thích ' giữ nguyên."""
    out, i, in_str = [], 0, False
    while i < len(line):
        ch = line[i]
        if in_str:
            out.append(ch)
            if ch == '"':
                in_str = False
        elif ch == '"':
            in_str = True
            out.append(ch)
        elif ch == "'":
            out.append(line[i:])
            break
        else:
            out.append(ch.lower())
        i += 1
    return "".join(out).rstrip()


def _norm(text: str) -> list:
    return [_canon(ln) for ln in text.replace("\r\n", "\n").split("\n") if ln.strip() != ""]


def info() -> None:
    raw = open(BAS_SRC, "rb").read()
    sha = hashlib.sha1(raw).hexdigest()
    project = f"msmcp_{sha[:8]}"
    mvba = os.path.join(WORK, project + ".mvba")
    print(f"Nguồn          : {BAS_SRC}")
    print(f"SHA-1 nguồn    : {sha}")
    print(f"SHA-256 nguồn  : {hashlib.sha256(raw).hexdigest()}")
    print(f"Chỉ ASCII      : {all(b < 128 for b in raw)}")
    print(f"Project VBA    : {project}")
    print(f"File project   : {mvba} ({'đã có' if os.path.exists(mvba) else 'chưa dựng'})")
    print("Các key-in MicroStation dùng để dựng (lần đầu):")
    print(f"  vba create {mvba}")
    print(f'  vba execute VBE.VBProjects("{project}").VBComponents.Import "{BAS_COPY}"')
    print(f"  vba save {project}")
    print("Lần sau chỉ: vba load <file project>;  mỗi lệnh đọc/ghi: vba run [" + project + "]McpFast.Main")


def build() -> None:
    mode = engine.ensure()
    print(f"Bộ máy: {mode}" + (f" (lỗi VBA: {engine.last_error})" if engine.last_error else ""))
    if mode == "vba":
        print(f"Project đã nạp: {engine.project}")


def verify() -> int:
    if engine.ensure() != "vba":
        print(f"Không nạp được VBA: {engine.last_error}")
        return 2
    out = os.path.join(tempfile.gettempdir(), f"msmcp_export_{int(time.time())}.bas")
    com.run(lambda: com.app().CadInputQueue.SendKeyin(
        f'vba execute VBE.VBProjects("{engine.project}").VBComponents("McpFast").Export "{out}"'))
    for _ in range(100):
        if os.path.exists(out) and os.path.getsize(out) > 0:
            break
        time.sleep(0.1)
    else:
        print("Không xuất được module từ MicroStation")
        return 2
    time.sleep(0.3)
    loaded = _norm(open(out, encoding="ascii", errors="replace").read())
    source = _norm(open(BAS_SRC, encoding="ascii").read())
    os.remove(out)
    diff = list(difflib.unified_diff(source, loaded, "vba/McpFast.bas (repo)", "McpFast (đang chạy)", lineterm=""))
    if diff:
        print("KHÁC NHAU:")
        print("\n".join(diff[:200]))
        return 1
    print(f"KHỚP: module McpFast trong project {engine.project} trùng {len(source)} dòng với vba/McpFast.bas")
    return 0


if __name__ == "__main__":
    for s in (sys.stdout, sys.stderr):
        try:
            s.reconfigure(encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass
    cmd = sys.argv[1] if len(sys.argv) > 1 else "info"
    if cmd == "info":
        info()
    elif cmd == "build":
        build()
    elif cmd == "verify":
        sys.exit(verify())
    else:
        print(__doc__)
        sys.exit(2)
