"""QThread 工作线程 — 非阻塞任务执行"""

from __future__ import annotations

from PyQt6.QtCore import QThread, pyqtSignal

from pyqt6.app.models.task_model import TaskItem, TaskStatus
from pyqt6.app.core.adapter import execute_task


class TaskWorker(QThread):
    """在后台线程运行单个任务"""

    # 信号定义
    task_started = pyqtSignal(str)           # task_id
    task_progress = pyqtSignal(str, int)     # task_id, percent
    task_log = pyqtSignal(str, str)          # task_id, message
    task_finished = pyqtSignal(str, bool, str)  # task_id, ok, message

    def __init__(self, item: TaskItem, parent=None):
        super().__init__(parent)
        self.item = item
        self._cancelled = False

    def run(self):
        tid = self.item.task_id
        self.item.status = TaskStatus.RUNNING
        self.task_started.emit(tid)
        self.task_log.emit(tid, f"开始任务: {self.item.task_type} → {self.item.input_path}")

        try:
            resp = execute_task(self.item)
            if self._cancelled:
                self.item.status = TaskStatus.CANCELLED
                self.task_finished.emit(tid, False, "任务已取消")
                return

            if resp.ok:
                self.item.status = TaskStatus.SUCCESS
                summary = resp.summary
                msg = f"完成: 处理 {summary.processed}, 失败 {summary.failed}, 跳过 {summary.skipped}"
                self.task_log.emit(tid, msg)
                self.task_progress.emit(tid, 100)
                self.task_finished.emit(tid, True, msg)
            else:
                self.item.status = TaskStatus.FAILED
                err = resp.error or {}
                msg = f"失败: [{err.get('code', '')}] {err.get('message', '未知错误')}"
                self.task_log.emit(tid, msg)
                self.task_finished.emit(tid, False, msg)

        except Exception as e:
            self.item.status = TaskStatus.FAILED
            msg = f"异常: {type(e).__name__}: {e}"
            self.task_log.emit(tid, msg)
            self.task_finished.emit(tid, False, msg)

    def cancel(self):
        self._cancelled = True
