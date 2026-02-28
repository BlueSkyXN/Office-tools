"""后台线程执行器 — 通过 root.after() 回调更新 UI"""

from __future__ import annotations

import threading
import traceback
from typing import Any, Callable

import sys
from pathlib import Path

_DOCX_ROOT = str(Path(__file__).resolve().parent.parent.parent.parent)
if _DOCX_ROOT not in sys.path:
    sys.path.insert(0, _DOCX_ROOT)

from core.api import TaskRequest, TaskResponse, run_task  # noqa: E402


class TaskWorker:
    """在后台线程中执行 core.run_task，完成后通过 root.after() 回调主线程"""

    def __init__(self, root: Any):
        """
        Args:
            root: Tk root window, 用于 .after() 调度
        """
        self._root = root
        self._thread: threading.Thread | None = None
        self._cancel_event = threading.Event()

    @property
    def running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def start(
        self,
        request: TaskRequest,
        on_progress: Callable[[str], None] | None = None,
        on_done: Callable[[TaskResponse], None] | None = None,
        on_error: Callable[[str], None] | None = None,
    ) -> None:
        """启动后台任务"""
        if self.running:
            return
        self._cancel_event.clear()
        self._thread = threading.Thread(
            target=self._run,
            args=(request, on_progress, on_done, on_error),
            daemon=True,
        )
        self._thread.start()

    def cancel(self) -> None:
        self._cancel_event.set()

    def _run(
        self,
        request: TaskRequest,
        on_progress: Callable[[str], None] | None,
        on_done: Callable[[TaskResponse], None] | None,
        on_error: Callable[[str], None] | None,
    ) -> None:
        try:
            if on_progress:
                self._schedule(on_progress, f"开始处理: {request.input_path}")

            response = run_task(request)

            if self._cancel_event.is_set():
                if on_progress:
                    self._schedule(on_progress, "任务已取消")
                return

            if on_done:
                self._schedule(on_done, response)
        except Exception as e:
            tb = traceback.format_exc()
            if on_error:
                self._schedule(on_error, f"{e}\n{tb}")

    def _schedule(self, callback: Callable, *args: Any) -> None:
        """线程安全地在主线程执行回调"""
        self._root.after(0, callback, *args)
