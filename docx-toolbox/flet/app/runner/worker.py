"""后台线程 Worker — 通过 Flet 线程安全机制更新 UI"""

from __future__ import annotations

import threading
import time
from typing import Callable

from core.api import TaskRequest, TaskResponse
from core.runner import TaskRunner, Job, JobStatus
from app.core.adapter import build_request, execute
from app.state.app_state import AppState, TaskState


class Worker:
    """在后台线程中执行单文件任务"""

    def __init__(self, state: AppState, page_update: Callable[[], None]) -> None:
        self._state = state
        self._page_update = page_update
        self._thread: threading.Thread | None = None

    def run_single(
        self,
        task_key: str,
        task_type: str,
        input_path: str,
        output_dir: str | None,
        options: dict,
        workers: int = 1,
        on_done: Callable[[TaskResponse], None] | None = None,
    ) -> None:
        ts = self._state.get_task(task_key)
        if ts.running:
            return
        ts.running = True
        ts.cancelled = False
        ts.progress = 0.0

        def _work():
            try:
                req = build_request(task_type, input_path, output_dir, options, workers)
                self._state.add_log(f"[{task_type}] 开始处理: {input_path}")
                self._page_update()

                resp = execute(req)

                ts.progress = 1.0
                if resp.ok:
                    summary = resp.summary
                    self._state.add_log(
                        f"[{task_type}] 完成: 处理 {summary.processed}, "
                        f"失败 {summary.failed}, 跳过 {summary.skipped}"
                    )
                else:
                    err = resp.error or {}
                    self._state.add_log(f"[{task_type}] 失败: {err.get('message', '未知错误')}")
            except Exception as e:
                self._state.add_log(f"[{task_type}] 异常: {e}")
            finally:
                ts.running = False
                self._page_update()
                if on_done:
                    on_done(resp if "resp" in dir() else None)

        self._thread = threading.Thread(target=_work, daemon=True)
        self._thread.start()


class BatchWorker:
    """使用 core.runner.TaskRunner 执行批量任务"""

    def __init__(self, state: AppState, page_update: Callable[[], None]) -> None:
        self._state = state
        self._page_update = page_update
        self._runner: TaskRunner | None = None
        self._thread: threading.Thread | None = None

    def run_batch(
        self,
        task_type: str,
        file_paths: list[str],
        output_dir: str | None,
        options: dict,
        workers: int = 1,
        on_done: Callable[[list[Job]], None] | None = None,
    ) -> None:
        ts = self._state.get_task("batch")
        if ts.running:
            return
        ts.running = True
        ts.cancelled = False
        ts.progress = 0.0

        self._runner = TaskRunner(max_workers=workers)

        for fp in file_paths:
            req = build_request(task_type, fp, output_dir, options, workers=1)
            self._runner.submit(req)

        def _on_progress(job: Job, current: int, total: int):
            ts.progress = current / max(total, 1)
            status = job.status.value
            self._state.add_log(f"[批量] {job.request.input_path} → {status} ({current}/{total})")
            self._page_update()

        self._runner.set_progress_callback(_on_progress)

        def _work():
            try:
                jobs = self._runner.run_all()
                ts.progress = 1.0
                success = sum(1 for j in jobs if j.status == JobStatus.SUCCESS)
                failed = sum(1 for j in jobs if j.status == JobStatus.FAILED)
                self._state.add_log(f"[批量] 全部完成: 成功 {success}, 失败 {failed}")
            except Exception as e:
                self._state.add_log(f"[批量] 异常: {e}")
            finally:
                ts.running = False
                self._page_update()
                if on_done:
                    on_done(self._runner.jobs if self._runner else [])

        self._thread = threading.Thread(target=_work, daemon=True)
        self._thread.start()

    def cancel(self) -> None:
        if self._runner:
            self._runner.cancel()
            self._state.get_task("batch").cancelled = True
