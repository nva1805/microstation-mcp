"""
src/core/com.py
Kết nối COM tới MicroStation V8i.

Khác bản gốc:
- Chỉ bám vào phiên MicroStation ĐANG CHẠY (GetActiveObject). Không bao giờ tự
  khởi động ustation.exe, không ghi registry.
- Mọi lời gọi COM chạy trên MỘT luồng STA riêng (hàng đợi tuần tự), nên nhiều tool
  gọi song song không giẫm lên nhau.
"""

from __future__ import annotations

import queue
import threading
from concurrent.futures import Future
from typing import Any, Callable, TypeVar

import pythoncom
import win32com.client

T = TypeVar("T")

PROGID = "MicroStationDGN.Application"


class MicroStationNotRunning(RuntimeError):
    pass


class NoActiveFile(RuntimeError):
    pass


class ComWorker:
    """Một luồng STA duy nhất giữ đối tượng Application và thực thi mọi lời gọi COM."""

    def __init__(self) -> None:
        self._q: "queue.Queue[tuple[Callable[[], Any], Future]]" = queue.Queue()
        self._app = None
        self._thread = threading.Thread(target=self._loop, name="msmcp-com", daemon=True)
        self._thread.start()

    # -- luồng COM -----------------------------------------------------------
    def _loop(self) -> None:
        pythoncom.CoInitialize()
        while True:
            fn, fut = self._q.get()
            if fut.set_running_or_notify_cancel():
                try:
                    fut.set_result(fn())
                except BaseException as ex:  # noqa: BLE001 - chuyển lỗi về caller
                    fut.set_exception(ex)

    def _connect(self):
        try:
            app = win32com.client.GetActiveObject(PROGID)
        except pythoncom.com_error as ex:
            raise MicroStationNotRunning(
                "MicroStation V8i chưa chạy. Hãy mở MicroStation và một file .dgn rồi thử lại "
                "(MCP này không tự khởi động MicroStation)."
            ) from ex
        return app

    def app(self, require_file: bool = True):
        """Chỉ gọi BÊN TRONG luồng COM (tức là trong hàm truyền cho run())."""
        if self._app is not None:
            try:
                _ = self._app.Name  # còn sống?
            except Exception:  # noqa: BLE001
                self._app = None
        if self._app is None:
            self._app = self._connect()
        if require_file:
            try:
                has = bool(self._app.HasActiveDesignFile)
            except Exception:  # noqa: BLE001
                self._app = None
                raise MicroStationNotRunning("Mất kết nối tới MicroStation, hãy thử lại.")
            if not has:
                raise NoActiveFile("MicroStation đang chạy nhưng chưa mở file .dgn nào.")
        return self._app

    # -- API ------------------------------------------------------------------
    def run(self, fn: Callable[[], T], timeout: float = 600.0) -> T:
        """Chạy fn() trên luồng COM và chờ kết quả."""
        if threading.current_thread() is self._thread:
            return fn()
        fut: Future = Future()
        self._q.put((fn, fut))
        return fut.result(timeout=timeout)


com = ComWorker()
