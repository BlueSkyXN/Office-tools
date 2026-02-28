"""后台线程任务执行器"""

from __future__ import annotations

import threading
from typing import Callable

from core.api import TaskRequest, TaskResponse, run_task
from core.logging_utils import get_logger

logger = get_logger()


class BackgroundWorker:
    """在后台线程中执行 core.run_task，不阻塞 UI"""

    def __init__(self):
        self._threads: dict[str, threading.Thread] = {}
        self._cancel_flags: dict[str, threading.Event] = {}

    def submit(
        self,
        task_id: str,
        request: TaskRequest,
        on_complete: Callable[[TaskResponse], None],
    ):
        """提交任务到后台线程"""
        cancel_event = threading.Event()
        self._cancel_flags[task_id] = cancel_event

        def _cancelled_response(detail: str) -> TaskResponse:
            return TaskResponse(
                ok=False,
                task_id=task_id,
                status="failed",
                error={
                    "code": "E_CANCELLED",
                    "message": "任务被用户取消",
                    "detail": detail,
                },
            )

        def _run():
            try:
                if cancel_event.is_set():
                    on_complete(_cancelled_response("cancelled before start"))
                    return
                response = run_task(request, cancel_event=cancel_event)
                if cancel_event.is_set() and response.ok:
                    response = _cancelled_response("cancelled during execution")
                on_complete(response)
            except Exception as e:
                if cancel_event.is_set():
                    on_complete(_cancelled_response(f"cancelled with exception: {e}"))
                    return
                logger.error("后台任务异常 task=%s: %s", task_id, e)
                error_response = TaskResponse(
                    ok=False,
                    task_id=task_id,
                    status="failed",
                    error={
                        "code": "E_INTERNAL",
                        "message": str(e),
                        "detail": type(e).__name__,
                    },
                )
                on_complete(error_response)
            finally:
                self._threads.pop(task_id, None)
                self._cancel_flags.pop(task_id, None)

        thread = threading.Thread(target=_run, name=f"task-{task_id}", daemon=True)
        self._threads[task_id] = thread
        thread.start()

    def cancel(self, task_id: str):
        """设置取消标志"""
        event = self._cancel_flags.get(task_id)
        if event:
            event.set()
