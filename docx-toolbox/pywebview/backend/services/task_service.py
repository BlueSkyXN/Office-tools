"""任务服务层 — 封装 core.run_task 的异步执行"""

from __future__ import annotations

import threading
import uuid
from datetime import datetime, timezone
from typing import Any

from core.api import TaskRequest, RuntimeOptions
from core.logging_utils import setup_file_logging

from backend.models.task_model import TaskRecord, TaskStatus
from backend.runner.worker import BackgroundWorker


class TaskService:
    """管理任务的创建、执行、取消与查询"""

    def __init__(self):
        self._lock = threading.Lock()
        self._tasks: dict[str, TaskRecord] = {}
        self._worker = BackgroundWorker()

    def start(self, payload: dict) -> dict:
        """创建并启动后台任务，返回任务信息"""
        task_id = uuid.uuid4().hex[:12]
        now = datetime.now(timezone.utc).isoformat()

        record = TaskRecord(
            task_id=task_id,
            task_type=payload["task_type"],
            input_path=payload["input_path"],
            output_dir=payload.get("output_dir"),
            options=payload.get("options", {}),
            status=TaskStatus.PENDING,
            created_at=now,
        )

        log_path = setup_file_logging(task_id=task_id)
        record.log_path = str(log_path)

        with self._lock:
            record.status = TaskStatus.RUNNING
            self._tasks[task_id] = record

        runtime_opts = payload.get("runtime", {})
        request = TaskRequest(
            task_id=task_id,
            task_type=record.task_type,
            input_path=record.input_path,
            output_dir=record.output_dir,
            options=record.options,
            runtime=RuntimeOptions(
                workers=runtime_opts.get("workers", 1),
                dry_run=runtime_opts.get("dry_run", False),
            ),
        )

        def on_complete(response):
            resp = response.to_dict()
            err = resp.get("error") or {}
            err_code = str(err.get("code", ""))
            with self._lock:
                current = self._tasks.get(task_id)
                if current is None:
                    return
                if current.status == TaskStatus.CANCELLED:
                    # 保持取消状态，避免被后台回调覆盖。
                    if err_code == "E_CANCELLED":
                        current.error = err or None
                    return

                if err_code == "E_CANCELLED":
                    current.status = TaskStatus.CANCELLED
                    current.error = err or None
                    return

                if response.ok:
                    current.status = TaskStatus.SUCCESS
                    current.summary = resp.get("summary")
                    current.error = None
                else:
                    current.status = TaskStatus.FAILED
                    current.error = err or None
        self._worker.submit(task_id, request, on_complete)
        with self._lock:
            return self._tasks[task_id].to_dict()

    def cancel(self, task_id: str):
        """取消指定任务"""
        should_cancel_worker = False
        with self._lock:
            record = self._tasks.get(task_id)
            if record is None:
                raise KeyError(task_id)
            if record.status in {TaskStatus.SUCCESS, TaskStatus.FAILED, TaskStatus.CANCELLED}:
                return
            record.status = TaskStatus.CANCELLED
            should_cancel_worker = True

        if should_cancel_worker:
            self._worker.cancel(task_id)

    def get_status(self, task_id: str) -> dict:
        """查询任务状态"""
        with self._lock:
            record = self._tasks.get(task_id)
            if record is None:
                raise KeyError(task_id)
            return record.to_dict()

    def list_all(self) -> list[dict]:
        """列出所有任务（最近的在前）"""
        with self._lock:
            return [r.to_dict() for r in reversed(self._tasks.values())]

    def get_log_path(self, task_id: str) -> str | None:
        """返回指定任务的日志文件路径"""
        with self._lock:
            record = self._tasks.get(task_id)
            if record is None:
                raise KeyError(task_id)
            return record.log_path
