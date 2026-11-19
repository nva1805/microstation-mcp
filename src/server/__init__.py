"""
src/server/__init__.py
Tạo FastMCP server và đăng ký công cụ.
"""

from fastmcp import FastMCP

from src.tools import register

SERVER_NAME = "microstation-v8i"

INSTRUCTIONS = """\
MicroStation V8i MCP - đọc/ghi bản đồ địa chính và KCS theo TT 26/2024 (PL21, PL22).

Cách làm việc:
- Đọc nhanh bằng bộ máy VBA chạy trong MicroStation (đọc ~27.000 đối tượng trong 1-2 giây);
  nếu VBA không dùng được thì tự chuyển COM (chậm).
- ms_summary trước để nắm bản vẽ; ms_query/ms_get để xem chi tiết; kcs_check để kiểm tra.
- Muốn đọc file khác mà KHÔNG đổi file người dùng đang xem: truyền `file` (mở ngầm chỉ đọc).
- Ghi (ms_edit) chỉ trên file đang mở; luôn chạy dry_run=True trước, cho người dùng xem,
  rồi mới dry_run=False. Có sao lưu tự động và ms_undo cho cả lô.
- ms_view zoom_ids/zoom_to + ảnh chụp để cho người dùng thấy vị trí lỗi.
- Level dùng SỐ (code) theo PL21: 10 ranh thửa hiện trạng, 13 số thửa, 2 loại đất, 4 diện tích,
  14 tường nhà, 23 chỉ giới đường, 32 kênh mương, 46 địa giới xã, 61 ranh theo giấy tờ, 63 khung.
"""


def create_server() -> FastMCP:
    server = FastMCP(name=SERVER_NAME, instructions=INSTRUCTIONS)
    register(server)
    return server


__all__ = ["create_server"]
