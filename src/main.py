"""MicroStation V8i MCP - điểm khởi chạy (stdio). Log chỉ ra stderr để không làm hỏng giao thức."""

import logging
import os
import sys

for _s in (sys.stdin, sys.stdout, sys.stderr):
    _s.reconfigure(encoding="utf-8", errors="replace")

logging.basicConfig(stream=sys.stderr, level=os.environ.get("MSMCP_LOG", "WARNING"),
                    format="%(asctime)s %(levelname)s %(name)s: %(message)s")

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.server import create_server  # noqa: E402

def main() -> None:
    create_server().run(transport="stdio", show_banner=False)


if __name__ == "__main__":
    main()
