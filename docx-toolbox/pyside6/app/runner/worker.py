"""QThread workers calling core.runner.TaskRunner"""

from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import QThread, Signal

_DOCX_TOOLBOX_ROOT = str(Path(__file__).resolve().parent.parent.parent.parent)
if _DOCX_TOOLBOX_ROOT not in sys.path:
    sys.path.insert(0, _DOCX_TOOLBOX_ROOT)

from core.api import TaskRequest, TaskResponse, run_task  # noqa: E402
from core.runner import TaskRunner, Job, JobStatus  # noqa: E402


class TaskWorker(QThread):
    """Runs a single TaskRequest on a background thread."""
    finished = Signal(object)    # TaskResponse
    error = Signal(str)
    log_message = Signal(str)

    def __init__(self, request: TaskRequest, parent=None):
        super().__init__(parent)
        self._request = request
        self._cancelled = False

    def run(self):
        try:
            self.log_message.emit(f"开始任务: {self._request.task_type} — {self._request.input_path}")
            response: TaskResponse = run_task(self._request)
            if self._cancelled:
                return
            self.finished.emit(response)
            if response.ok:
                self.log_message.emit(f"任务完成: 处理 {response.summary.processed} 个文件")
            else:
                err_msg = response.error.get("message", "未知错误") if response.error else "未知错误"
                self.error.emit(err_msg)
                self.log_message.emit(f"任务失败: {err_msg}")
        except Exception as e:
            if not self._cancelled:
                self.error.emit(str(e))
                self.log_message.emit(f"任务异常: {e}")

    def cancel(self):
        self._cancelled = True


class BatchWorker(QThread):
    """Runs multiple tasks via core.runner.TaskRunner."""
    job_updated = Signal(int, str)   # (job_index, status_string)
    progress = Signal(int, int)      # (current, total)
    all_finished = Signal(list)      # list[Job]
    log_message = Signal(str)

    def __init__(self, requests: list[TaskRequest], max_workers: int = 1, parent=None):
        super().__init__(parent)
        self._requests = requests
        self._max_workers = max_workers
        self._runner: TaskRunner | None = None

    def run(self):
        self._runner = TaskRunner(max_workers=self._max_workers)
        for req in self._requests:
            self._runner.submit(req)

        self._runner.set_progress_callback(self._on_progress)
        self.log_message.emit(f"批量任务开始: {len(self._requests)} 个任务")
        jobs = self._runner.run_all()
        self.all_finished.emit(jobs)
        success = sum(1 for j in jobs if j.status == JobStatus.SUCCESS)
        failed = sum(1 for j in jobs if j.status == JobStatus.FAILED)
        self.log_message.emit(f"批量任务完成: 成功 {success}, 失败 {failed}")

    def cancel(self):
        if self._runner:
            self._runner.cancel()

    def _on_progress(self, job: Job, current: int, total: int):
        idx = self._runner.jobs.index(job) if job in self._runner.jobs else -1
        self.job_updated.emit(idx, job.status.value)
        self.progress.emit(current, total)
